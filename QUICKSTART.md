# 快速启动指南

欢迎使用 DOCX 解析器！这个指南将帮助您在几分钟内开始使用。

## 🚀 一键启动演示

```bash
# 克隆项目并启动演示
git clone https://github.com/SansAppel-PW/Docx_parser.git
cd Docx_parser
python demo.py
```

## 📋 系统要求

- Python 3.7 或更高版本
- 推荐使用虚拟环境

## ⚡ 快速安装

### 最小化安装（基本功能）
```bash
pip install python-docx lxml Pillow
```

### 完整安装（推荐）
```bash
pip install python-docx lxml Pillow opencv-python Wand
```

### 系统依赖（可选，用于高级图片处理）
```bash
# macOS
brew install imagemagick

# Ubuntu/Debian
sudo apt-get install imagemagick libmagickwand-dev

# Windows - 下载安装 ImageMagick
# https://imagemagick.org/script/download.php#windows
```

## 🎯 使用方式

### 1. 演示模式（推荐新手）
```bash
python demo.py
```

### 2. 直接使用解析器
```python
from src.parsers.document_parser import parse_docx

# 解析文档
result = parse_docx("your_document.docx", "output_folder")
print(f"提取了 {len(result.get('images', {}))} 张图片")
```

### 3. 命令行模式
```bash
# 处理单个文件
python docx_parser_modular.py Files/examples/demo.docx

# 批量处理
python docx_parser_modular.py input_folder
```

## 📁 输出结果

解析完成后，您将得到：
- `📄 JSON 文件`: 结构化的文档内容
- `🖼️ images/ 文件夹`: 提取的图片
- `📊 详细报告`: 处理过程和统计信息

## 🔧 故障排除

**问题：导入模块失败**
```bash
# 确保在项目根目录运行
cd Docx_parser
python demo.py
```

**问题：ImageMagick 相关错误**
```bash
# 安装系统依赖后重试
pip install Wand
```

**问题：文档无法解析**
```bash
# 尝试快速模式
# 在代码中添加 quick_mode=True 参数
```

## 📖 更多信息

- 📚 [详细文档](README.md)
- 🔧 [使用手册](USAGE.md)
- 🐛 [问题报告](https://github.com/SansAppel-PW/Docx_parser/issues)
- 💬 [讨论交流](https://github.com/SansAppel-PW/Docx_parser/discussions)

---

🎉 **开始您的文档解析之旅吧！**
