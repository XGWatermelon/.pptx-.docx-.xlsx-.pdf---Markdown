#!/usr/bin/env python3
"""
免费高效的文档转换器（Docling + watchdog）
支持: .pptx .docx .xlsx .pdf  ->  Markdown
特性: 增量处理（哈希记录） + 自动监听新建文件
     每个文档独立文件夹 + 图片存放在 assets + Obsidian wiki 图片语法
"""
import hashlib
import pickle
import re
import time
from pathlib import Path
from typing import Optional
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from docling.document_converter import DocumentConverter

# ========== 配置区域 ==========
SOURCE_DIR = Path(r"F:\personal_files\AI练习\SOURCE_DIR")
OUTPUT_DIR = Path(r"F:\personal_files\AI练习\output")
STATE_FILE = Path(".conversion_state.pkl")
SUPPORTED_EXT = {".pptx", ".docx", ".xlsx", ".pdf"}
WATCH_MODE = False  # True: 转换完后继续监听新文件; False: 转换完后退出
# ==============================

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def get_file_hash(filepath: Path) -> str:
    """计算文件 SHA256 哈希"""
    hasher = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            hasher.update(chunk)
    return hasher.hexdigest()

def load_state():
    if STATE_FILE.exists():
        with open(STATE_FILE, "rb") as f:
            return pickle.load(f)
    return {}

def save_state(state):
    with open(STATE_FILE, "wb") as f:
        pickle.dump(state, f)

def preprocess_docling_md(md_content: str) -> str:
    """清理 Docling 导出的 markdown 格式干扰（粗体标记、HTML 实体等）"""
    # 解码 HTML 实体
    md_content = md_content.replace('&amp;', '&')
    md_content = md_content.replace('&lt;', '<')
    md_content = md_content.replace('&gt;', '>')
    # 去除 Docling 误加的粗体标记（**...** 或 **...*）
    lines = md_content.split('\n')
    result = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('**') and len(stripped) > 2:
            # 去掉开头的 **
            inner = stripped[2:]
            # 如果末尾也有 ** 也去掉
            if inner.endswith('**'):
                inner = inner[:-2]
            result.append(inner.strip())
            continue
        result.append(line)
    return '\n'.join(result)

def extract_code_from_table_line(line: str, domain: str) -> tuple:
    """检查表格行的单元格中是否包含代码，如果有则提取出来。
    返回 (是否修改, 提取的代码行列表, 修改后的表格行)"""
    if not line.strip().startswith('|'):
        return False, [], line

    cells = line.split('|')
    # 至少需要3个元素（前导空、内容、尾部空）
    if len(cells) < 3:
        return False, [], line

    code_markers = {
        'abap': [
            r'REPORT\s+z', r'DATA:\s*\w+', r'DATA\s+\w+\s+TYPE',
            r'TYPES:\s', r'TYPES\s+BEGIN\s+OF', r'CALL\s+FUNCTION',
            r'^\s*IF\s+', r'^\s*LOOP\s+AT', r'^\s*SELECT\s+', r'^\s*WRITE',
        ],
    }
    patterns = code_markers.get(domain, [])
    if not patterns:
        return False, [], line

    extracted_code = []
    modified = False
    for i, cell in enumerate(cells):
        cell_content = cell.strip()
        if not cell_content:
            continue
        # 检查单元格是否包含代码
        code_count = sum(1 for p in patterns if re.search(p, cell_content, re.IGNORECASE))
        if code_count >= 2:  # 至少匹配2个代码模式才认为是代码
            # 按两个以上连续空格分割，提取代码行
            code_lines = re.split(r'\s{2,}', cell_content)
            code_lines = [cl.strip() for cl in code_lines if cl.strip()]
            if len(code_lines) >= 3:  # 至少3行才认为是代码块
                extracted_code.extend(code_lines)
                cells[i] = ''  # 清空单元格
                modified = True

    if modified:
        new_line = '|'.join(cells)
        return True, extracted_code, new_line
    return False, [], line

def detect_code_blocks(md_content: str) -> str:
    """扫描段落关键词，驱动代码块语言标签"""

    # 先清理 Docling 格式干扰
    md_content = preprocess_docling_md(md_content)

    # 领域关键词 → 语言标签
    domain_keywords = {
        'abap': [
            # ===== 核心技术关键词 =====
            r'\bABAP\b', r'\bSAP\b', r'\bS/4HANA\b', r'\bECC\b', r'\bR/3\b',
            r'BAPI(?:[\\_]|[_\b])', r'RFC(?:[\\_]|[_\b])', r'\bIDoc\b', r'\bALE\b',
            r'\bBAdI\b', r'\bBADI\b', r'\bUser\s*Exit\b', r'\bEnhancement\b',
            r'\bBAPI_TRANSACTION\b', r'\bCALL\s+FUNCTION\b', r'\bRFC_CALL\b',
            r'\bFiori\b', r'\bOData\b', r'\bCDS\s*View\b', r'\bAMDP\b',
            r'\bBOPF\b', r'\bRAP\b', r'\bCAP\b',
            # ===== MM 物料管理 =====
            r'\bME[0-9]', r'\bME5[0-9]', r'\bMIGO\b', r'\bMIRO\b', r'\bMIR7\b',
            r'\bMB0[0-9]', r'\bMB5[0-9]', r'\bMM0[0-9]', r'\bMMBE\b', r'\bMM60\b',
            r'\bMK0[0-9]', r'\bXK0[0-9]',
            r'\b采购信息记录\b', r'\b采购订单\b', r'\b采购申请\b', r'\b采购发票\b',
            r'\b物料凭证\b', r'\b物料主数据\b', r'\b供应商主数据\b', r'\b收货\b', r'\b发货\b',
            r'\bBAPI_PO_CREATE\b', r'\bBAPI_GOODSMVT_CREATE\b', r'\bBAPI_MATERIAL_SAVEDATA\b',
            r'\bBAPI_REQUISITION_CREATE\b', r'\bBAPI_REQUISITION_CHANGE\b', r'\bBAPI_REQUISITION_RELEASE\b',
            r'\bBAPI_INCOMINGINVOICE\b', r'\bME_INFORECORD\b',
            # ===== SD 销售与分销 =====
            r'\bVA0[0-9]', r'\bVA1[0-9]', r'\bVA2[0-9]',
            r'\bVF0[0-9]', r'\bVF1[0-9]', r'\bVF2[0-9]',
            r'\bVL0[0-9]', r'\bVL1[0-9]', r'\bVL2[0-9]',
            r'\bVK1[0-9]', r'\bXD0[0-9]',
            r'\b销售订单\b', r'\b交货单\b', r'\b发票\b', r'\b定价\b', r'\b客户主数据\b',
            r'\bBAPI_SALESORDER\b', r'\bBAPI_BILLINGDOC\b', r'\bBAPI_OUTB_DELIVERY\b',
            r'\bBAPI_CUSTOMER\b', r'\bSD_SALESDOCUMENT\b',
            # ===== FI 财务会计 =====
            r'\bFB0[0-9]', r'\bFB5[0-9]', r'\bF-0[0-9]', r'\bF-1[0-9]', r'\bF-2[0-9]', r'\bF-3[0-9]', r'\bF-4[0-9]',
            r'\bFK0[0-9]', r'\bFK1[0-9]', r'\bFK2[0-9]', r'\bFK3[0-9]',
            r'\bFS1[0-9]', r'\bFAGL\b', r'\bFBL1N\b', r'\bFBL5N\b',
            r'\b会计凭证\b', r'\b总账\b', r'\b科目\b', r'\b过账\b', r'\b清账\b',
            r'\bBAPI_ACC_DOCUMENT\b', r'\bBAPI_GL\b', r'\bBAPI_AP_ACC\b', r'\bBAPI_AR_ACC\b',
            r'\bBAPI_COMPANYCODE\b',
            # ===== CO 管理会计 =====
            r'\bKS0[0-9]', r'\bKA0[0-9]', r'\bKE5[0-9]', r'\bKB2[0-9]',
            r'\bCJ20N\b', r'\bCJ0[0-9]',
            r'\b成本中心\b', r'\b成本要素\b', r'\b利润中心\b', r'\b内部订单\b',
            r'\bBAPI_COSTCENTER\b', r'\bBAPI_COSTELEMENT\b', r'\bBAPI_INTERNALORDER\b',
            r'\bBAPI_COSTACTPLN\b',
            # ===== PP 生产计划 =====
            r'\bCO0[0-9]', r'\bMD0[0-9]', r'\bMD1[0-9]', r'\bMD2[0-9]',
            r'\bCS0[0-9]', r'\bCR0[0-9]', r'\bCR1[0-9]',
            r'\b生产订单\b', r'\bBOM\b', r'\b工作中心\b', r'\b工艺路线\b', r'\bMRP\b',
            r'\bBAPI_PRODORD\b', r'\bBAPI_MATERIAL_BOM\b',
            # ===== PM 设备维护 =====
            r'\bIW2[0-9]', r'\bIW3[0-9]', r'\bIW6[0-9]',
            r'\b工单\b', r'\b通知单\b', r'\b设备主数据\b',
            # ===== QM 质量管理 =====
            r'\bQA0[0-9]', r'\bQA1[0-9]', r'\bQA2[0-9]',
            r'\b检验批\b', r'\b质量通知\b',
            # ===== HR 人力资源 =====
            r'\bPA20\b', r'\bPA30\b', r'\bPA40\b', r'\bPB[0-9]',
            r'\b信息类型\b', r'\b人事事件\b',
            # ===== 通用 ABAP 关键字 =====
            r'\bSELECT\b.*\bFROM\b', r'\bINSERT\b.*\bINTO\b', r'\bUPDATE\b.*\bSET\b',
            r'\bMODIFY\b', r'\bDELETE\b.*\bFROM\b',
            r'\bLOOP\s+AT\b', r'\bENDLOOP\b', r'\bREAD\s+TABLE\b',
            r'\bAPPEND\b', r'\bSORT\b', r'\bDELETE\s+ADJACENT\b',
            r'\bCLASS\b.*\bDEFINITION\b', r'\bCLASS\b.*\bIMPLEMENTATION\b',
            r'\bINTERFACE\b', r'\bMETHOD\b', r'\bENDMETHOD\b',
            r'\bFORM\b', r'\bENDFORM\b', r'\bPERFORM\b',
            r'\bMODULE\b', r'\bENDMODULE\b', r'\bSCREEN\b',
            r'\bAT\s+SELECTION-SCREEN\b', r'\bSTART-OF-SELECTION\b', r'\bEND-OF-SELECTION\b',
            r'\bTOP-OF-PAGE\b', r'\bEND-OF-PAGE\b',
        ],
        'xml': [r'\bXML\b', r'\bxmlns\b', r'\bWSDL\b', r'\bSOAP\b', r'\bXSD\b', r'\bSchema\b',
                r'\bwsdl:\b', r'\bsoap:\b', r'\bxs:\b'],
        'json': [r'\bJSON\b', r'\bobject\b', r'\barray\b', r'\bkey-value\b'],
        'python': [r'\bPython\b', r'\bpip\b', r'\bimport\s+\w+', r'\bdef\s+\w+', r'\bclass\s+\w+'],
        'groovy': [r'\bGroovy\b', r'\bCPI\b', r'\biFlow\b', r'\bscript\b', r'\bcom\.sap\b'],
    }

    # 代码起始模式（按领域）
    code_starters = {
        'abap': [
            # ===== 程序声明 =====
            r'^REPORT\s+', r'^FUNCTION-POOL\b', r'^TYPE-POOL\b',
            # ===== 数据声明 =====
            r'^DATA:?\s*$', r'^DATA:\s*\w+', r'^DATA\s+\w+',
            r'^CONSTANTS:?\s*$', r'^CONSTANTS\s+\w+', r'^CONSTANT\s+\w+',
            r'^TYPES:?\s*$', r'^TYPES\s+BEGIN\s+OF', r'^TYPES\s+\w+',
            r'^STATICS:?\s*$', r'^STATICS\s+\w+',
            r'^CLASS-DATA\b', r'^CLASS-METHODS\b', r'^CLASS-EVENTS\b',
            # ===== 结构体/内表 =====
            r'^BEGIN\s+OF\b', r'^END\s+OF\b',
            r'^STANDARD\s+TABLE\b', r'^SORTED\s+TABLE\b', r'^HASHED\s+TABLE\b',
            # ===== 字段符号与引用 =====
            r'^FIELD-SYMBOLS\b', r'^FIELD-SYMBOL\b',
            r'^ASSIGN\b', r'^UNASSIGN\b',
            # ===== 选择屏幕 =====
            r'^PARAMETERS\s*:', r'^SELECT-OPTIONS\s*:',
            r'^SELECTION-SCREEN\b', r'^AT\s+SELECTION-SCREEN\b',
            r'^INITIALIZATION\b', r'^AT\s+USER-COMMAND\b',
            # ===== 事件块 =====
            r'^START-OF-SELECTION\b', r'^END-OF-SELECTION\b',
            r'^TOP-OF-PAGE\b', r'^END-OF-PAGE\b',
            r'^AT\s+LINE-SELECTION\b', r'^AT\s+NEW\b', r'^AT\s+END\b',
            r'^GET\s+\w+', r'^GET\s+LATE\b',
            # ===== 函数/子程序调用 =====
            r'^CALL\s+FUNCTION\s+', r'^CALL\s+METHOD\b', r'^CALL\s+SCREEN\b',
            r'^CALL\s+DIALOG\b', r'^CALL\s+TRANSACTION\b',
            r'^SUBMIT\b', r'^LEAVE\s+TO\b', r'^LEAVE\s+LIST-PROCESSING\b',
            # ===== 子程序/方法定义 =====
            r'^FORM\s+\w+', r'^ENDFORM\b',
            r'^METHOD\s+\w+', r'^ENDMETHOD\b',
            r'^MODULE\s+\w+', r'^ENDMODULE\b',
            r'^CLASS\s+\w+\s+DEFINITION\b', r'^CLASS\s+\w+\s+IMPLEMENTATION\b',
            r'^INTERFACE\s+\w+', r'^ENDINTERFACE\b',
            # ===== 数据库操作 =====
            r'^SELECT\b', r'^INSERT\b', r'^UPDATE\b', r'^MODIFY\b', r'^DELETE\b',
            r'^OPEN\s+DATASET\b', r'^CLOSE\s+DATASET\b', r'^TRANSFER\b', r'^READ\s+DATASET\b',
            # ===== 控制流 =====
            r'^IF\s+', r'^ELSEIF\s+', r'^ELSE\b', r'^ENDIF\b',
            r'^CASE\s+', r'^WHEN\s+', r'^ENDCASE\b',
            r'^DO\s+', r'^ENDDO\b', r'^WHILE\s+', r'^ENDWHILE\b',
            r'^LOOP\s+AT\b', r'^ENDLOOP\b',
            r'^READ\s+TABLE\b', r'^CHECK\b', r'^EXIT\b', r'^CONTINUE\b', r'^RETURN\b',
            r'^TRY\b', r'^CATCH\b', r'^CLEANUP\b', r'^ENDTRY\b',
            # ===== 输出 =====
            r'^WRITE:\s*/', r'^WRITE\s+/', r'^WRITE\b',
            r'^NEW-LINE\b', r'^SKIP\b', r'^ULINE\b', r'^POSITION\b',
            r'^FORMAT\b', r'^HIDE\b',
            # ===== 内表操作 =====
            r'^APPEND\b', r'^INSERT\s+TABLE\b', r'^COLLECT\b',
            r'^SORT\b', r'^DELETE\s+ADJACENT\b', r'^DELETE\s+TABLE\b',
            r'^MODIFY\s+TABLE\b', r'^CLEAR\b', r'^FREE\b', r'^REFRESH\b',
            r'^MOVE\b', r'^MOVE-CORRESPONDING\b', r'^CORRESPONDING\b',
            # ===== 字符串操作 =====
            r'^CONCATENATE\b', r'^SPLIT\b', r'^CONDENSE\b', r'^TRANSLATE\b',
            r'^REPLACE\b', r'^FIND\b', r'^SHIFT\b', r'^STRLEN\b',
            # ===== 异常处理 =====
            r'^RAISE\b', r'^RAISE\s+EXCEPTION\b', r'^MESSAGE\b',
            r'^CATCH\s+SYSTEM-EXCEPTIONS\b', r'^ENDCATCH\b',
            # ===== ALV =====
            r'^REUSE_ALV\b', r'^CL_SALV\b', r'^CL_GUI_ALV\b',
            # ===== BAPI 调用 =====
            r'^BAPI_TRANSACTION_COMMIT\b', r'^BAPI_TRANSACTION_ROLLBACK\b',
            # ===== 变量命名约定 =====
            r'^lv_?\w+', r'^ls_?\w+', r'^lt_?\w+', r'^gv_?\w+', r'^gs_?\w+', r'^gt_?\w+',
            r'^rv_?\w+', r'^rs_?\w+', r'^rt_?\w+', r'^cv_?\w+', r'^cs_?\w+', r'^ct_?\w+',
            r'^ev_?\w+', r'^es_?\w+', r'^et_?\w+', r'^iv_?\w+', r'^is_?\w+', r'^it_?\w+',
            r'^mv_?\w+', r'^ms_?\w+', r'^mt_?\w+', r'^mo_?\w+',
            r'^<\w+>',  # 字段符号 <fs_xxx>
            # ===== 注释与分隔线 =====
            r'^&-{5,}', r'^\*&', r'^\*\s', r'^" ',
        ],
        'xml': [
            r'^<\?xml', r'^<\w+:definitions', r'^<\w+:types', r'^<\w+:message',
            r'^<\w+:portType', r'^<\w+:binding', r'^<\w+:service',
            r'^<xs:schema', r'^<xs:element', r'^<xs:complexType',
        ],
        'json': [
            r'^\{', r'^\[',
        ],
        'python': [
            r'^import\s+', r'^from\s+\w+\s+import', r'^def\s+\w+', r'^class\s+\w+',
            r'^if\s+__name__\s*==\s*__main__', r'^print\(',
        ],
        'groovy': [
            r'^import\s+', r'^def\s+\w+', r'^class\s+\w+',
            r'^//', r'^/\*',
        ],
    }

    lines = md_content.split('\n')
    result_lines = []
    in_code_block = False
    current_domain = None  # 当前检测到的领域

    # 预扫描：统计全文领域关键词出现次数，找到主导领域
    domain_scores = {d: 0 for d in domain_keywords}
    for line in lines:
        for domain, patterns in domain_keywords.items():
            for pattern in patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    domain_scores[domain] += 1
    dominant_domain = max(domain_scores, key=domain_scores.get) if any(domain_scores.values()) else None

    def is_table_line(line):
        stripped = line.strip()
        return stripped.startswith('|') or stripped.startswith('---')

    def detect_domain_from_text(line):
        """从文本行检测领域关键词"""
        for domain, patterns in domain_keywords.items():
            for pattern in patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    return domain
        return None

    def is_code_start(line, domain):
        """判断是否为对应领域的代码起始"""
        if domain not in code_starters:
            return False
        for pattern in code_starters[domain]:
            if re.search(pattern, line, re.IGNORECASE):
                return True
        return False

    def is_any_code_start(line):
        """判断是否为任意领域的代码起始"""
        for domain in code_starters:
            if is_code_start(line, domain):
                return domain
        return None

    for line in lines:
        # 已在代码块中
        if in_code_block:
            if is_table_line(line):
                in_code_block = False
                result_lines.append('```')
                result_lines.append(line)
            else:
                result_lines.append(line)
            continue

        # 未在代码块中

        # 0. 检查表格行中是否嵌入了代码
        if line.strip().startswith('|'):
            use_domain = current_domain or dominant_domain or 'abap'
            extracted, code_lines, new_line = extract_code_from_table_line(line, use_domain)
            if extracted:
                result_lines.append(new_line)
                result_lines.append(f'```{use_domain}')
                result_lines.extend(code_lines)
                result_lines.append('```')
                current_domain = use_domain
                continue

        # 1. 先检测段落中的领域关键词
        detected = detect_domain_from_text(line)
        if detected:
            current_domain = detected

        # 2. 检测是否为代码起始（先用当前领域检测，再用任意领域检测）
        code_domain = None
        if current_domain and is_code_start(line, current_domain):
            code_domain = current_domain
        else:
            code_domain = is_any_code_start(line)
            # 如果检测到代码起始但没有领域上下文，用预扫描的主导领域
            if code_domain and not current_domain:
                current_domain = dominant_domain or code_domain

        if code_domain:
            in_code_block = True
            result_lines.append(f'```{code_domain}')
            result_lines.append(line)
        else:
            result_lines.append(line)

    # 文件以代码块结束
    if in_code_block:
        result_lines.append('```')

    return '\n'.join(result_lines)

def convert_to_markdown(src_path: Path, out_dir: Path, converter: DocumentConverter) -> Optional[Path]:
    """使用 Docling 转换单个文档为 Markdown，图片存放在 assets 文件夹"""
    try:
        result = converter.convert(src_path)

        # 创建文档专属文件夹
        doc_name = src_path.stem
        doc_folder = out_dir / doc_name
        doc_folder.mkdir(parents=True, exist_ok=True)

        # 创建 assets 文件夹存放图片
        assets_folder = doc_folder / "assets"
        assets_folder.mkdir(parents=True, exist_ok=True)

        # 导出文档内容
        doc = result.document

        # 使用 docling 的导出功能，获取 markdown 和图片
        md_content = doc.export_to_markdown()

        # 处理图片：将图片保存到 assets 文件夹并更新引用为 Obsidian wiki 格式
        image_counter = 0

        # 遍历文档中的图片（doc.pictures）
        for pic in doc.pictures:
            img_ref = pic.image
            if img_ref and hasattr(img_ref, 'pil_image') and img_ref.pil_image:
                image_counter += 1
                image_filename = f"image_{image_counter:03d}.png"
                image_path = assets_folder / image_filename

                try:
                    img_ref.pil_image.save(str(image_path))
                except Exception as e:
                    print(f"  [!] 保存图片失败: {e}")

        # 替换 markdown 中的 <!-- image --> 占位符为 Obsidian wiki 格式
        def replace_image_placeholder(_):
            nonlocal image_counter
            replace_image_placeholder.counter += 1
            return f"![[image_{replace_image_placeholder.counter:03d}.png]]"

        replace_image_placeholder.counter = 0
        md_content = re.sub(r'<!-- image -->', replace_image_placeholder, md_content)

        # 格式化代码块（根据段落关键词自动识别语言）
        md_content = detect_code_blocks(md_content)

        # 写入 md 文件
        md_filename = doc_name + ".md"
        md_path = doc_folder / md_filename
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        return md_path
    except Exception as e:
        print(f"[X] 转换失败 {src_path.name}: {e}")
        import traceback
        traceback.print_exc()
        return None

def process_file(src_path: Path, converter: DocumentConverter) -> bool:
    """增量处理：仅当文件新增或内容变化时转换"""
    if src_path.suffix.lower() not in SUPPORTED_EXT:
        return False

    state = load_state()
    current_hash = get_file_hash(src_path)

    # 已处理且未修改 → 跳过
    if state.get(str(src_path)) == current_hash:
        print(f"[=] 跳过（未变化）: {src_path.name}")
        return True

    print(f"[*] 转换中: {src_path.name}")
    out_path = convert_to_markdown(src_path, OUTPUT_DIR, converter)
    if out_path:
        state[str(src_path)] = current_hash
        save_state(state)
        print(f"[+] 已生成: {out_path}")
        return True
    return False

class ConversionHandler(FileSystemEventHandler):
    def __init__(self, converter: DocumentConverter):
        self.converter = converter

    def on_created(self, event):
        if not event.is_directory:
            process_file(Path(event.src_path), self.converter)

def batch_convert_existing(converter: DocumentConverter):
    """启动监听前，先处理文件夹内所有已有文件"""
    print("[...] 扫描已有文件...")
    for src_path in SOURCE_DIR.rglob("*"):
        if src_path.is_file() and src_path.suffix.lower() in SUPPORTED_EXT:
            process_file(src_path, converter)

def start_watching(converter: DocumentConverter):
    """启动文件监听守护进程"""
    print(f"[...] 开始监听目录: {SOURCE_DIR}")
    handler = ConversionHandler(converter)
    observer = Observer()
    observer.schedule(handler, str(SOURCE_DIR), recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    converter = DocumentConverter()
    batch_convert_existing(converter)
    if WATCH_MODE:
        start_watching(converter)
    else:
        print("[OK] 全部转换完成")
