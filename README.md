# 文档压缩工具

一个强大的文档压缩和文本提取工具,支持多种文档格式,包括Word、PDF、PPT、Excel和图片文件。

## 功能特点

- 支持多种文档格式:
  - Word文档 (.docx, .doc)
  - PDF文档 (.pdf)
  - PowerPoint演示文稿 (.pptx, .ppt)
  - Excel工作簿 (.xlsx, .xls)
  - 纯文本文件 (.txt)
  - 图片文件 (.png, .jpg, .jpeg, .bmp, .tiff, .gif)

- 智能文本提取:
  - 自动识别文档类型
  - 提取文本内容
  - OCR识别图片中的文字
  - 表格数据提取

- 高效压缩:
  - 智能去重
  - 保留段落格式
  - 自动过滤无用信息
  - 多进程并行处理

## 系统要求

- Windows 10 或更高版本
- Python 3.8 或更高版本
- 至少 4GB 内存
- 至少 2GB 可用磁盘空间

## 安装步骤

1. 克隆或下载本项目

2. 安装Python依赖:
   ```bash
   pip install -r requirements.txt
   ```

3. 安装额外依赖:

   - Tesseract OCR:
     1. 下载地址: https://github.com/UB-Mannheim/tesseract/wiki
     2. 选择最新版本的Windows安装包
     3. 安装到默认目录: C:\Program Files\Tesseract-OCR
     4. 安装时选择"Additional language data (download)",确保选中"Chinese simplified"
     5. 将安装目录添加到系统环境变量PATH

   - Poppler (用于PDF处理):
     1. 下载地址: https://github.com/oschwartz10612/poppler-windows/releases/
     2. 解压到任意目录
     3. 将bin目录添加到系统环境变量PATH

## 使用方法

1. 准备文件:
   - 创建 raw 目录(如果不存在)
   - 将需要处理的文件放入 raw 目录

2. 运行程序:
   ```bash
   python compress_docs_stable.py
   ```

3. 等待处理完成:
   - 处理后的文件将保存在 output 目录
   - 每个文件会生成对应的 _compressed.txt 文件
   - 按 Ctrl+C 可随时中断处理

4. 查看结果:
   - 程序会显示每个文件的处理结果
   - 包括原始大小、压缩大小和压缩比
   - 最后会显示总体统计信息

## 注意事项

1. 文件准备:
   - 确保文件没有被其他程序占用
   - 不要处理正在编辑的文件
   - 建议处理前备份重要文件

2. 系统资源:
   - 程序会自动调整使用的进程数
   - 处理大文件时可能需要较多内存
   - 建议关闭不必要的程序

3. 特殊情况:
   - 某些加密或受保护的文件可能无法处理
   - 图片质量过低可能影响OCR识别效果
   - 特殊格式的表格可能无法完整提取

## 常见问题

1. OCR识别率低:
   - 检查是否正确安装Tesseract
   - 确认是否安装中文语言包
   - 尝试提供更清晰的图片

2. PDF处理失败:
   - 检查是否正确安装Poppler
   - 确认PDF文件未加密
   - 尝试使用其他PDF查看器打开

3. Office文件处理失败:
   - 确认已安装相应的Office组件
   - 检查文件是否被占用
   - 尝试另存为新文件再处理

## 更新日志

### v1.0 (2024-03-22)
- 初始版本发布
- 支持多种文档格式
- 实现基本的压缩功能
- 添加OCR支持
- 多进程并行处理

## 许可证

MIT License 
