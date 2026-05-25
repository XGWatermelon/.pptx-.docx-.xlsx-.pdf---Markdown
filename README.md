# Document to Markdown Converter

基于 Docling 的文档转换工具，将 `.pptx` `.docx` `.xlsx` `.pdf` 转换为 Markdown 格式，针对 SAP/ABAP 领域文档做了专项优化。

## 特性

- 支持 `.pptx` `.docx` `.xlsx` `.pdf` 四种格式
- 领域驱动的代码块自动检测（ABAP、XML、JSON、Python、Groovy）
- Obsidian wiki 图片语法 `![[image.png]]`
- 每个文档独立文件夹，图片存放在 `assets` 子目录
- SHA256 哈希增量处理，跳过未变化文件
- 可选的文件监听模式（watchdog）

## 安装

```bash
pip install docling watchdog
```

## 使用

1. 修改 `convert_documents.py` 顶部的配置区域：

```python
SOURCE_DIR = Path(r"你的源文件目录")
OUTPUT_DIR = Path(r"你的输出目录")
WATCH_MODE = False  # True: 转换后继续监听新文件
```

2. 运行：

```bash
python convert_documents.py
```

## 输出结构

```
output/
  文档名称/
    文档名称.md
    assets/
      image_001.png
      image_002.png
```

## 代码块检测

转换器会扫描文档中的领域关键词（如 ABAP、SAP、BAPI 等），当检测到对应领域的代码起始模式时，自动将代码包裹在对应语言的代码块中：

```
```abap
CALL FUNCTION 'BAPI_INCOMINGINVOICE_CANCEL'
  EXPORTING
    invoicedocnumber = lv_inv_doc_no
    ...
```
```
