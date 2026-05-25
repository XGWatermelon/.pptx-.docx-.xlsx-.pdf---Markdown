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
        'abap': [r'\bABAP\b', r'\bSAP\b', r'BAPI(?:[\\_]|[_\b])', r'RFC(?:[\\_]|[_\b])', r'\bS/4HANA\b', r'\bECC\b',
                 r'\bME[0-9]', r'\bVA[0-9]', r'\bVF[0-9]', r'\bMIGO\b', r'\bSE[0-9]',
                 r'\b采购信息记录\b', r'\b采购订单\b', r'\b采购发票\b', r'\b物料凭证\b'],
        'xml': [r'\bXML\b', r'\bxmlns\b', r'\bWSDL\b', r'\bSOAP\b', r'\bXSD\b', r'\bSchema\b',
                r'\bwsdl:\b', r'\bsoap:\b', r'\bxs:\b'],
        'json': [r'\bJSON\b', r'\bobject\b', r'\barray\b', r'\bkey-value\b'],
        'python': [r'\bPython\b', r'\bpip\b', r'\bimport\s+\w+', r'\bdef\s+\w+', r'\bclass\s+\w+'],
        'groovy': [r'\bGroovy\b', r'\bCPI\b', r'\biFlow\b', r'\bscript\b', r'\bcom\.sap\b'],
    }

    # 代码起始模式（按领域）
    code_starters = {
        'abap': [
            r'^REPORT\s+[zy]', r'^DATA:?\s*$', r'^DATA:\s*\w+', r'^DATA\s+\w+', r'^TYPES:?\s*$',
            r'^TYPES\s+BEGIN\s+OF', r'^CALL\s+FUNCTION\s+', r'^WRITE:\s*/',
            r'^WRITE\s+/', r'^PARAMETERS\s*:', r'^SELECT-OPTIONS\s*:',
            r'^START-OF-SELECTION', r'^END-OF-SELECTION', r'^FORM\s+\w+',
            r'^MODULE\s+\w+', r'^CLASS\s+\w+\s+DEFINITION',
            r'^lv_?\w+', r'^ls_?\w+', r'^lt_?\w+', r'^gv_?\w+', r'^gs_?\w+', r'^gt_?\w+',
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
