# DOCX 解析器

[![Pyt## 📁 项目结构

```
├── src/                    # 核心源代码（模块化架构）
│   ├── extractors/         # 内容提取器（图片、SmartArt、嵌入对象）
│   ├── parsers/           # 文档解析器（结构解析、批量处理）
│   ├── processors/        # 文本处理器（标准化输出）
│   └── utils/             # 工具函数（文本处理、图片处理等）
├── examples/              # 使用示例和演示脚本
├── tests/                 # 测试和分析工具
├── docs/                  # 详细文档和指南
├── legacy/                # 历史版本备份
├── Files/examples/        # 示例DOCX文档
├── parsed_docs/examples/  # 示例解析结果
└── docx_parser_modular.py # 主程序入口
```g.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

功能强大的 Word 文档解析工具，支持提取文本、图片、SmartArt、嵌入对象等内容。

## ✨ 特性

- 🚀 **高性能模块化架构** - 清晰的代码组织和可扩展设计
- 📝 **完整内容提取** - 文本、表格、列表、标题层次结构
- 🖼️ **智能图片处理** - PNG、JPEG、SVG等格式，支持绝对路径
- 🎨 **SmartArt图表解析** - 完整的SmartArt内容和结构提取
- 📎 **嵌入对象支持** - Visio、Excel等Office嵌入对象
- 📋 **标准化文本输出** - 符合规范的结构化文本格式
- ⚡ **批量处理模式** - 高效处理大量文档
- 🔄 **智能去重机制** - 基于内容哈希的文件去重

## � 项目结构

```
├── src/                    # 核心源代码（模块化架构）
│   ├── extractors/         # 内容提取器
│   ├── parsers/           # 文档解析器
│   ├── processors/        # 文本处理器
│   └── utils/             # 工具函数
├── examples/              # 使用示例和演示脚本
├── tests/                 # 测试和分析工具
├── docs/                  # 详细文档
├── legacy/                # 历史版本备份
├── Files/examples/        # 示例文档
└── docx_parser_modular.py # 主程序入口
```

## �🛠️ 安装

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
# 基础功能演示
python examples/demo.py

# 完整文本处理流程演示
python examples/demo_text_processing.py
```

### 命令行使用
```bash
# 使用默认路径（处理Files/examples/demo.docx）
python docx_parser_modular.py

# 处理单个文件
python docx_parser_modular.py Files/examples/demo.docx

# 批量处理文件夹
python docx_parser_modular.py Files/PLM2.0

# 指定自定义输出目录
python docx_parser_modular.py Files/examples/demo.docx my_output
```

### 代码调用
```python
from src.parsers.document_parser import parse_docx
from src.processors.text_processor import process_document_to_text

# 基础解析
result = parse_docx(
    docx_path="Files/examples/demo.docx",
    output_dir="output",
    quick_mode=True  # 推荐使用快速模式
)

# 生成标准化文本
text_output = process_document_to_text(
    result, 
    output_dir="output/demo"
)
```

## 📊 输出格式

解析完成后，每个文档会生成以下文件结构：

### 📁 输出目录结构
```
output_directory/
├── document.json          # 完整的JSON结构化数据
├── processed_text.txt     # 标准化文本文件
├── images/               # 图片文件目录
│   ├── img_[hash].png
│   └── img_[hash].jpg
├── smartart/             # SmartArt数据目录
│   └── smartart_[hash].json
└── embedded_objects/     # 嵌入对象目录
    ├── embedded_obj_[hash].json
    └── preview_[hash].emf
```

### 📋 JSON 结构化数据 (`document.json`)
```json
{
  "metadata": {
    "source_path": "path/to/document.docx",
    "title": "文档标题",
    "author": "作者姓名",
    "created": "2024-01-01T00:00:00+00:00",
    "modified": "2024-01-01T00:00:00+00:00",
    "file_size": "XXX.XX KB"
  },
  "sections": [
    {
      "type": "section",
      "title": "章节标题",
      "level": 1,
      "content": [
        {
          "type": "paragraph",
          "text": "段落内容",
          "bold": false,
          "italic": false
        },
        {
          "type": "table",
          "index": 0,
          "rows": [
            ["表头1", "表头2", "表头3"],
            ["数据1", "数据2", "数据3"]
          ]
        }
      ]
    }
  ],
  "images": {
    "img_[hash]": {
      "filename": "img_[hash].png",
      "format": "PNG",
      "path": "/absolute/path/to/image.png",
      "width": 800,
      "height": 600
    }
  },
  "processing_info": {
    "total_time": "0.5s",
    "warnings": [],
    "errors": []
  }
}
```

### 📝 标准化文本文档 (`processed_text.txt`)

基于实际解析结果的格式示例：
```text
# 文档标题
## 一级标题  
## 二级标题
### 三级标题

文档名称-章节序号 章节标题：
<|SECTION|>
    文档名称-章节序号 章节标题-子章节序号 子章节标题：
    <|PARAGRAPH|>段落文本内容...</|PARAGRAPH|>
</|SECTION|>

<|SECTION|>
    文档名称-章节序号 章节标题-子章节序号 子章节标题：
    <|LISTITEM|>列表项内容</|LISTITEM|>
    <|IMAGE|>![image](/absolute/path/to/image.png)</|IMAGE|>
    <|TABLE|>
        <|ROW|>|列标题1|列标题2|列标题3|</|ROW|>
        <|ROW|>|数据1|数据2|数据3|</|ROW|>
    </|TABLE|>
</|SECTION|>
```

### 🏷️ 标签格式规范
- **文档结构**: 使用标准的Markdown标题格式 (`#`, `##`, `###`)
- **章节内容**: 使用 `<|SECTION|></|SECTION|>` 包装结构化章节
- **段落文本**: 使用 `<|PARAGRAPH|></|PARAGRAPH|>` 包装段落内容
- **列表项**: 使用 `<|LISTITEM|></|LISTITEM|>` 包装列表项目
- **图片引用**: 使用 `<|IMAGE|>![alt](path)</|IMAGE|>` 格式，支持绝对路径
- **表格数据**: 使用 `<|TABLE|><|ROW|>|cell1|cell2|</|ROW|></|TABLE|>` 格式

### 📁 资源文件管理
- **images/**: 提取的图片文件，使用内容哈希命名防重复
- **smartart/**: SmartArt图表的JSON结构化数据
- **embedded_objects/**: Office嵌入对象及其预览图

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

## � 文档

- [详细使用指南](docs/usage-guide.md) - 完整的使用说明和API文档
- [文本处理器指南](docs/text-processor-guide.md) - 文本处理模块详细说明
- [更新日志](CHANGELOG.md) - 版本更新记录
- [需求规格](REQUIREMENTS.md) - 功能需求和技术规格

## �🔧 故障排除

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
- [查看文档](docs/usage-guide.md)

## 📈 版本历史

### v2.0.0 (2025-08-13) - 重大重构版本
- ✨ **新特性**
  - 完全模块化架构重构
  - 添加文件内容去重机制，防止重复存储
  - 实现标准化标签格式 `<|TAG|></|TAG|>`
  - 支持图片绝对路径生成
  - 新增灵活的命令行接口，支持默认路径处理

- 🔧 **改进**
  - 项目结构完全重组，模块化设计
  - 文档系统全面升级，分层次管理
  - 性能优化，处理速度提升
  - 代码质量改进，添加类型提示和详细注释

- 🗂️ **结构变更**
  - 创建 `src/` 目录，模块化核心代码
  - 新增 `examples/` 目录，统一管理示例
  - 新增 `docs/` 目录，集中文档管理
  - 新增 `legacy/` 目录，保存历史版本
  - 优化 `.gitignore`，只保留示例相关内容

- 🐛 **修复**
  - 修复表格格式处理问题
  - 解决文件命名冲突导致的重复存储
  - 改进错误处理和日志记录

---

⭐ 如果项目对您有帮助，请给个 Star！
