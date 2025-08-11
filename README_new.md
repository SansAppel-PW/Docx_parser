# DOCX 解析器

[![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

功能强大的 Word 文档解析工具，支持提取文本、图片、SmartArt、嵌入对象等内容。

## ✨ 特性

- 🚀 高性能模块化架构
- 📝 完整内容提取（文本、表格、列表）
- 🖼️ 智能图片处理（PNG、JPEG、SVG 等）
- 🎨 SmartArt 图表解析
- 📎 嵌入对象支持（Visio、Excel）
- ⚡ 快速模式和批量处理

## 🛠️ 安装

```bash
# 基础安装
pip install python-docx lxml Pillow

# 完整功能（推荐）
pip install python-docx lxml Pillow opencv-python Wand

# 系统依赖（可选，用于高级图片处理）
# macOS: brew install imagemagick
# Linux: sudo apt-get install imagemagick libmagickwand-dev
```

## 🚀 快速开始

### 演示脚本（推荐）
```bash
python demo.py
```

### 命令行使用
```bash
# 单个文件
python docx_parser_modular.py Files/examples/demo.docx

# 批量处理
python docx_parser_modular.py input_folder
```

### 代码调用
```python
from src.parsers.document_parser import parse_docx

result = parse_docx("document.docx", "output_dir")
print(f"提取图片: {len(result.get('images', {}))}")
```

## 📁 项目结构

```
Docx_parser/
├── src/                    # 核心模块
│   ├── extractors/        # 内容提取器
│   ├── parsers/           # 文档解析器
│   └── utils/             # 工具函数
├── Files/examples/        # 示例文档
├── parsed_docs/examples/  # 解析结果
├── demo.py               # 演示脚本
└── docx_parser_modular.py # 主程序
```

## 📊 输出格式

```json
{
  "metadata": {
    "title": "文档标题",
    "author": "作者"
  },
  "sections": [
    {
      "type": "paragraph",
      "text": "段落内容"
    }
  ],
  "images": {
    "img_001": {
      "filename": "img_001.png",
      "format": "PNG",
      "width": 800,
      "height": 600
    }
  },
  "processing_info": {
    "total_time": "处理时间",
    "warnings": [],
    "errors": []
  }
}
```

## ⚙️ 配置

### 快速模式
```python
result = parse_docx("document.docx", "output", quick_mode=True)
```

### 批量处理
```python
from src.parsers.batch_processor import process_docx_folder

count = process_docx_folder("docs", "results", quick_mode=True)
```

## 🔧 故障排除

| 问题 | 解决方案 |
|------|----------|
| 导入模块失败 | 确保在项目根目录运行 |
| ImageMagick 错误 | 安装系统依赖：`brew install imagemagick` |
| 内存不足 | 使用快速模式：`quick_mode=True` |

## 🤝 贡献

1. Fork 仓库
2. 创建分支 (`git checkout -b feature/NewFeature`)
3. 提交更改 (`git commit -m 'Add NewFeature'`)
4. 推送分支 (`git push origin feature/NewFeature`)
5. 创建 Pull Request

## 📄 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件

## 📞 支持

- [报告问题](https://github.com/SansAppel-PW/Docx_parser/issues)
- [功能建议](https://github.com/SansAppel-PW/Docx_parser/issues)
- [查看文档](USAGE.md)

---

⭐ 如果项目对您有帮助，请给个 Star！
