# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

AI 练习项目集合。当前包含一个文档转换工具。

## 依赖

```bash
pip install docling watchdog
```

## 运行文档转换器

```bash
# 修改 convert_documents.py 中的配置区域后再运行
python convert_documents.py
```

配置项（`convert_documents.py` 顶部）：
- `SOURCE_DIR`：源文件夹路径
- `OUTPUT_DIR`：Markdown 输出目录
- `SUPPORTED_EXT`：支持的扩展名（默认 `.pptx .docx .xlsx .pdf`）

## 架构

`convert_documents.py` 是单文件脚本，流程：
1. 启动时批量扫描 `SOURCE_DIR` 中已有文件并转换
2. 通过 watchdog 监听新建文件，实时触发转换
3. 使用 SHA256 哈希 + pickle 状态文件实现增量处理（跳过未变化文件）
4. `DocumentConverter` 实例在 `__main__` 创建一次，传入各函数复用
