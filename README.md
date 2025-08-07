# DOCX解析器

高性能的Word文档解析工具，支持提取文本、图片、SmartArt和嵌入对象。

## 快速开始

```bash
# 处理单个文件
python docx_parser.py Files/example.docx

# 批量处理文件夹
python docx_parser.py Files/PLM2.0

# 快速测试
python quick_parse_example.py
```

## 项目结构

- `Files/` - 存放原始Word文档
- `parsed_docs/` - 存放解析结果
- `Test/` - 测试脚本和开发工具

详细说明请查看 `项目结构说明.md`

## 特性

- 🚀 快速模式：跳过耗时的图像转换
- 📊 结构化输出：JSON格式的文档结构
- 🖼️ 图片提取：支持多种图片格式
- 🎨 SmartArt解析：提取图表文本内容
- 📦 嵌入对象：检测Visio、Excel等嵌入内容
- 📈 批量处理：支持文件夹批量处理

## 性能

- 单个文档处理时间：< 1秒
- 相比原版性能提升：97%+
- 支持大批量文档处理
