# 文档中心

欢迎来到DOCX解析器的文档中心！这里包含了使用、开发和贡献本项目所需的全部信息。

## 📚 文档导航

### 🚀 快速开始
- [项目主页](../README.md) - 项目概述和快速开始指南
- [安装指南](#安装指南) - 详细的安装步骤和依赖配置
- [快速示例](#快速示例) - 5分钟快速上手

### 📖 用户文档
- [使用指南](usage-guide.md) - 完整的使用说明和最佳实践
- [命令行参考](#命令行参考) - 详细的CLI参数说明
- [Python API文档](#api文档) - 编程接口参考
- [配置选项](#配置选项) - 自定义配置指南

### 🔧 技术文档
- [文本处理器指南](text-processor-guide.md) - 文本处理模块的详细技术说明
- [架构设计](#架构设计) - 系统架构和模块设计
- [输出格式规范](#输出格式) - JSON和文本输出的详细格式
- [性能优化](#性能优化) - 性能调优和最佳实践

### 🛠️ 开发文档
- [需求规格说明](../REQUIREMENTS.md) - 完整的功能需求和技术规格
- [更新日志](../CHANGELOG.md) - 版本历史和变更记录
- [贡献指南](#贡献指南) - 如何参与项目开发
- [测试指南](#测试指南) - 测试框架和测试用例

---

## 🚀 安装指南

### 系统要求
- **Python版本**: 3.7 或更高版本
- **操作系统**: Windows 10+, macOS 10.14+, Linux (Ubuntu 18.04+)
- **内存**: 建议 4GB+ RAM
- **磁盘空间**: 至少 500MB 可用空间

### 快速安装
```bash
# 1. 克隆仓库
git clone https://github.com/SansAppel-PW/Docx_parser.git
cd Docx_parser

# 2. 安装基础依赖
pip install python-docx lxml Pillow

# 3. 运行测试
python docx_parser_modular.py
```

### 完整安装（推荐）
```bash
# 安装所有功能依赖
pip install python-docx lxml Pillow opencv-python

# 可选：高级图片处理（需要系统依赖）
# macOS
brew install imagemagick
pip install Wand

# Ubuntu/Debian
sudo apt-get install imagemagick libmagickwand-dev
pip install Wand

# Windows
# 下载并安装 ImageMagick，然后运行：
pip install Wand
```

---

## ⚡ 快速示例

### 处理单个文档
```bash
# 使用默认示例
python docx_parser_modular.py

# 处理指定文档
python docx_parser_modular.py my_document.docx
```

### Python API调用
```python
from src.parsers.document_parser import parse_docx
from src.processors.text_processor import process_document_to_text

# 解析文档
result = parse_docx("document.docx", "output", quick_mode=True)

# 生成标准化文本
if result:
    text_path = process_document_to_text(result, "output/document")
    print(f"文本文件: {text_path}")
```

---

## 🖥️ 命令行参考

### 基本语法
```bash
python docx_parser_modular.py [INPUT] [OUTPUT]
```

### 参数说明
- **INPUT**: 输入文件或文件夹路径（可选，默认使用示例文件）
- **OUTPUT**: 输出目录（可选，默认为 parsed_docs）

### 使用示例
```bash
# 无参数：处理默认示例文件
python docx_parser_modular.py

# 单个文件：处理指定文档
python docx_parser_modular.py document.docx

# 批量处理：处理整个文件夹
python docx_parser_modular.py documents_folder/

# 自定义输出：指定输出目录
python docx_parser_modular.py document.docx my_output/
```

---

## 📋 API文档

### 核心API

#### parse_docx()
```python
def parse_docx(docx_path, output_dir, quick_mode=True):
    """
    解析DOCX文档
    
    Args:
        docx_path (str): DOCX文件路径
        output_dir (str): 输出目录
        quick_mode (bool): 是否使用快速模式
        
    Returns:
        dict: 解析结果，包含文档内容和元数据
    """
```

#### process_document_to_text()
```python
def process_document_to_text(document_data, output_dir):
    """
    生成标准化文本文件
    
    Args:
        document_data (dict): 文档解析结果
        output_dir (str): 输出目录
        
    Returns:
        str: 生成的文本文件路径
    """
```

#### process_docx_folder()
```python
def process_docx_folder(input_folder, output_base_dir, quick_mode=True):
    """
    批量处理文档文件夹
    
    Args:
        input_folder (str): 输入文件夹路径
        output_base_dir (str): 输出基础目录
        quick_mode (bool): 是否使用快速模式
        
    Returns:
        int: 成功处理的文件数量
    """
```

---

## ⚙️ 配置选项

### 性能配置
```python
# 快速模式：提升处理速度，适合大批量处理
quick_mode = True

# 内存优化：减少内存使用，适合大文件处理
memory_efficient = True

# 并发处理：启用多进程处理（批量处理时）
max_workers = 4
```

### 输出配置
```python
# 输出格式配置
output_config = {
    "include_json": True,        # 生成JSON文件
    "include_text": True,        # 生成文本文件
    "include_images": True,      # 提取图片
    "include_smartart": True,    # 提取SmartArt
    "include_objects": True      # 提取嵌入对象
}
```

---

## 🏗️ 架构设计

### 模块架构
```
src/
├── extractors/     # 内容提取层
├── parsers/        # 文档解析层
├── processors/     # 文本处理层
└── utils/          # 工具函数层
```

### 数据流
```
DOCX文档 → 解析器 → 提取器 → 处理器 → 输出文件
```

### 设计原则
- **单一职责**: 每个模块专注特定功能
- **松耦合**: 模块间依赖最小化
- **可扩展**: 易于添加新功能
- **可测试**: 支持单元测试

---

## 📊 输出格式

### JSON结构
详细的JSON格式说明请参考 [使用指南](usage-guide.md#json格式)

### 文本格式
标准化文本格式使用统一的标签结构：
```text
<|PARAGRAPH|>段落内容</|PARAGRAPH|>
<|TABLE|>表格内容</|TABLE|>
<|IMAGE|>图片引用</|IMAGE|>
```

---

## 🚀 性能优化

### 处理大文件
- 启用快速模式：`quick_mode=True`
- 增加内存限制：调整系统内存分配
- 使用SSD存储：提高I/O性能

### 批量处理优化
- 并行处理：利用多核CPU
- 内存管理：及时清理临时文件
- 进度监控：实时反馈处理状态

---

## 🤝 贡献指南

### 开发环境设置
```bash
# 1. Fork仓库并克隆
git clone https://github.com/your-username/Docx_parser.git

# 2. 创建开发分支
git checkout -b feature/your-feature

# 3. 安装开发依赖
pip install -r requirements-dev.txt

# 4. 运行测试
python -m pytest tests/
```

### 代码规范
- 遵循PEP 8代码风格
- 添加类型提示
- 编写完整的文档字符串
- 保证测试覆盖率

### 提交流程
1. 编写代码和测试
2. 运行完整测试套件
3. 更新相关文档
4. 提交Pull Request

---

## 🧪 测试指南

### 运行测试
```bash
# 运行所有测试
python -m pytest tests/

# 运行特定测试
python tests/test_modular_parser.py

# 生成覆盖率报告
python -m pytest --cov=src tests/
```

### 测试类型
- **单元测试**: 测试单个函数/类
- **集成测试**: 测试模块间协作
- **端到端测试**: 测试完整处理流程
- **性能测试**: 测试处理速度和内存使用

---

## ❓ 常见问题

### Q: 如何处理大型文档？
A: 启用快速模式并确保足够的内存。对于超大文档，考虑分段处理。

### Q: 支持哪些图片格式？
A: 支持PNG、JPEG、GIF、BMP、TIFF、SVG、EMF、WMF等常见格式。

### Q: 如何自定义输出格式？
A: 可以修改文本处理器模块，或使用API自定义处理逻辑。

### Q: 批量处理时如何监控进度？
A: 程序会自动显示处理进度，并在完成后生成详细的统计报告。

---

## 📞 获取帮助

- 🐛 [报告问题](https://github.com/SansAppel-PW/Docx_parser/issues)
- 💡 [功能建议](https://github.com/SansAppel-PW/Docx_parser/discussions)
- 📧 邮件支持：contact@example.com
- 📚 查看[使用指南](usage-guide.md)获取更多详细信息

---

*文档最后更新：2025-08-13*
