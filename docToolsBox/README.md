# Word 处理工具箱

本项目包含两个用于批量处理 Word 文档的 Python 脚本，旨在提高文档格式转换和清理的效率。

## 目录

- [环境要求](#环境要求)
- [安装依赖](#安装依赖)
- [脚本功能说明](#脚本功能说明)
    - [1. doc2docx.py (批量 doc 转 docx)](#1-doc2docxpy-批量-doc-转-docx)
    - [2. remove_last_image.py (批量移除末尾图片)](#2-remove_last_imagepy-批量移除末尾图片)
- [注意事项](#注意事项)

## 环境要求

- **操作系统**: Windows (由于 `doc2docx.py` 依赖 COM 接口调用 Word)
- **软件**: Microsoft Word (必须安装，用于 `doc` 格式转换)
- **Python**: Python 3.6+

## 安装依赖

在使用脚本之前，请确保安装了必要的 Python 库：

```bash
pip install -r requirements.txt
pip install pywin32
```

> **注意**: `requirements.txt` 中包含 `python-docx`。`doc2docx.py` 需要额外的 `pywin32` 库来与 Microsoft Word 进行交互。

## 脚本功能说明

### 1. doc2docx.py (批量 doc 转 docx)

该脚本用于将当前目录及其子目录下的所有 `.doc` 文件批量转换为 `.docx` 文件。

**功能特点**:
- **递归扫描**: 自动查找当前目录及所有子文件夹中的 `.doc` 文件。
- **多线程处理**: 默认使用 4 个线程并行转换，提高处理速度。
- **断点续传**: 自动记录已处理的文件（`processed_files.txt`），中断后再次运行会跳过已处理文件。
- **自动清理**: 转换成功后会自动**删除**原始的 `.doc` 文件。
- **日志记录**:
    - `conversion_errors.log`: 记录转换失败的文件和错误详情。
    - `processed_files.txt`: 记录已成功处理的文件路径。

**使用方法**:
双击运行或在命令行中执行：
```bash
python doc2docx.py
```

### 2. remove_last_image.py (批量移除末尾图片)

该脚本用于批量扫描 `.docx` 文件，并尝试移除文档**最末尾**的一张图片。通常用于去除文档末尾的广告图片或签名图。

**功能特点**:
- **智能判断**: 脚本会从文档末尾向前扫描。
    - 如果先遇到**文本内容**，则认为文档以文本结尾，**不会**移除任何图片（避免误删正文插图）。
    - 只有在文档末尾（忽略空白段落）直接遇到图片时，才会将其移除。
- **多线程处理**: 默认使用 4 个线程并行处理。
- **断点续传**: 记录已处理文件（`processed_image_removal.txt`），避免重复处理。
- **日志记录**:
    - `image_removal_errors.log`: 记录处理出错的文件。
    - `processed_image_removal.txt`: 记录已扫描/处理过的文件。

**使用方法**:
双击运行或在命令行中执行：
```bash
python remove_last_image.py
```

## 注意事项

1.  **数据备份**: `doc2docx.py` 会在转换成功后**删除源文件**，`remove_last_image.py` 会直接**修改源文件**。建议在运行脚本前备份重要数据。
2.  **Word 干扰**: `doc2docx.py` 运行时会启动后台 Word 进程。请尽量避免在脚本运行时手动操作 Word，以免发生冲突。
3.  **隐藏文件**: 脚本会自动跳过以 `~$` 开头的临时文件。
