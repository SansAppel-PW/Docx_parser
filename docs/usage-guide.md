# 使用指南

本文档详细介绍如何使用DOCX解析器的各种功能。

## 🚀 快速入门

### 1. 环境准备

确保您的Python环境满足要求：

```bash
# 检查Python版本（需要3.7+）
python --version

# 安装必需依赖
pip install python-docx lxml Pillow

# 可选：安装增强功能依赖
pip install opencv-python Wand
```

### 2. 基本使用

#### 方法一：命令行使用

```bash
# 处理单个文件
python docx_parser_modular.py "Files/PLM2.0/BOM 审核申请.docx"

# 批量处理文件夹
python docx_parser_modular.py Files/PLM2.0

# 使用默认文件夹
python docx_parser_modular.py
```

#### 方法二：Python代码调用

```python
from src.parsers.document_parser import parse_docx

# 基本解析
result = parse_docx(
    docx_path="Files/example.docx",
    output_dir="output",
    quick_mode=True  # 推荐使用快速模式
)

# 检查结果
if result:
    print(f"✅ 解析成功")
    print(f"📄 总节点数: {len(result.get('sections', []))}")
    print(f"🖼️ 图片数量: {len(result.get('images', {}))}")
    print(f"🎨 SmartArt数量: {len(result.get('smartart', {}))}")
    print(f"📎 嵌入对象: {len(result.get('embedded_objects', {}))}")
else:
    print("❌ 解析失败")
```

## 📊 批量处理

### 处理整个文件夹

```python
from src.parsers.batch_processor import process_docx_folder

# 批量处理
processed_count = process_docx_folder(
    input_folder="Files/PLM2.0",
    output_folder="parsed_results",
    quick_mode=True  # 启用快速模式
)

print(f"成功处理 {processed_count} 个文件")
```

### 查看处理报告

批量处理完成后，会生成 `summary.json` 文件：

```python
import json

# 读取处理报告
with open("parsed_results/summary.json", "r", encoding="utf-8") as f:
    summary = json.load(f)

print(f"总文件数: {summary['total_files']}")
print(f"成功处理: {summary['processed']}")
print(f"失败文件: {summary['failed']}")
print(f"成功率: {summary['success_rate']}")
```

## 🎯 高级功能

### 1. 自定义输出目录结构

```python
import os
from src.parsers.document_parser import parse_docx

# 自定义输出目录
output_base = "custom_output"
doc_name = "报告文档"

# 创建规范的输出目录
output_dir = os.path.join(output_base, doc_name)
os.makedirs(output_dir, exist_ok=True)

# 解析文档
result = parse_docx("document.docx", output_dir, quick_mode=True)
```

### 2. 解析结果后处理

```python
import json

def analyze_document_content(result):
    """分析解析结果"""
    
    # 统计文本内容
    text_sections = [s for s in result.get('sections', []) 
                    if s.get('type') == 'paragraph' and s.get('text')]
    
    # 统计表格
    tables = [s for s in result.get('sections', []) 
             if s.get('type') == 'table']
    
    # 统计图片
    images = result.get('images', {})
    
    # 统计SmartArt
    smartarts = result.get('smartart', {})
    
    # 生成分析报告
    analysis = {
        "content_summary": {
            "text_paragraphs": len(text_sections),
            "tables": len(tables),
            "images": len(images),
            "smartart_charts": len(smartarts)
        },
        "text_preview": [s.get('text', '')[:100] for s in text_sections[:5]],
        "table_info": [{"rows": len(t.get('data', [])), 
                       "cols": len(t.get('data', [{}])[0]) if t.get('data') else 0} 
                      for t in tables],
        "image_formats": list(set(img.get('format', 'unknown') 
                                 for img in images.values()))
    }
    
    return analysis

# 使用示例
result = parse_docx("document.docx", "output")
if result:
    analysis = analyze_document_content(result)
    
    # 保存分析结果
    with open("output/analysis.json", "w", encoding="utf-8") as f:
        json.dump(analysis, f, ensure_ascii=False, indent=2)
```

### 3. 错误处理和调试

```python
import logging

# 设置详细日志
logging.basicConfig(level=logging.DEBUG)

def safe_parse_docx(docx_path, output_dir):
    """安全的文档解析函数"""
    try:
        result = parse_docx(docx_path, output_dir, quick_mode=True)
        
        if not result:
            print(f"❌ 解析失败: {docx_path}")
            return None
            
        # 检查警告和错误
        processing_info = result.get('processing_info', {})
        warnings = processing_info.get('warnings', [])
        errors = processing_info.get('errors', [])
        
        if warnings:
            print(f"⚠️ 发现 {len(warnings)} 个警告")
            for warning in warnings[:3]:  # 显示前3个警告
                print(f"   - {warning}")
                
        if errors:
            print(f"❌ 发现 {len(errors)} 个错误")
            for error in errors[:3]:  # 显示前3个错误
                print(f"   - {error}")
        
        return result
        
    except Exception as e:
        print(f"❌ 解析异常: {docx_path}")
        print(f"   错误详情: {str(e)}")
        return None

# 使用示例
result = safe_parse_docx("Files/example.docx", "output")
```

## 🔧 性能优化

### 1. 快速模式 vs 完整模式

```python
import time

def compare_modes(docx_path):
    """对比不同模式的性能"""
    
    # 快速模式
    start_time = time.time()
    quick_result = parse_docx(docx_path, "output_quick", quick_mode=True)
    quick_time = time.time() - start_time
    
    # 完整模式
    start_time = time.time()
    full_result = parse_docx(docx_path, "output_full", quick_mode=False)
    full_time = time.time() - start_time
    
    print(f"快速模式: {quick_time:.2f}秒")
    print(f"完整模式: {full_time:.2f}秒")
    print(f"性能提升: {(full_time/quick_time):.1f}x")

# 使用示例
compare_modes("Files/large_document.docx")
```

### 2. 内存管理

```python
import gc
import psutil
import os

def monitor_memory_usage(func, *args, **kwargs):
    """监控内存使用情况"""
    process = psutil.Process(os.getpid())
    
    # 解析前内存
    memory_before = process.memory_info().rss / 1024 / 1024  # MB
    
    # 执行解析
    result = func(*args, **kwargs)
    
    # 解析后内存
    memory_after = process.memory_info().rss / 1024 / 1024  # MB
    
    # 手动垃圾回收
    gc.collect()
    memory_final = process.memory_info().rss / 1024 / 1024  # MB
    
    print(f"解析前内存: {memory_before:.1f}MB")
    print(f"解析后内存: {memory_after:.1f}MB")
    print(f"回收后内存: {memory_final:.1f}MB")
    print(f"内存增长: {memory_after - memory_before:.1f}MB")
    
    return result

# 使用示例
result = monitor_memory_usage(
    parse_docx, 
    "Files/large_document.docx", 
    "output", 
    quick_mode=True
)
```

## 📝 输出结果说明

### 目录结构

```
output_dir/
├── document.json       # 主要解析结果
├── images/            # 提取的图片文件
│   ├── img_001.png
│   ├── img_002.jpg
│   └── ...
├── smartart/          # SmartArt图表JSON文件
│   ├── smartart_001.json
│   └── ...
├── embedded_objects/  # 嵌入对象
│   ├── embedded_obj_001.json
│   └── ...
└── files/            # 其他文件
```

### JSON结构详解

```json
{
  "metadata": {
    "title": "文档标题",
    "author": "文档作者", 
    "subject": "文档主题",
    "created": "2024-01-01T10:00:00",
    "modified": "2024-01-02T15:30:00",
    "pages": 10,
    "words": 5000,
    "paragraphs": 100
  },
  "sections": [
    {
      "type": "paragraph",
      "text": "段落文本内容",
      "level": 1,
      "style": "标题1",
      "alignment": "left",
      "content": []  // 子内容（图片、表格等）
    },
    {
      "type": "table", 
      "rows": 3,
      "columns": 4,
      "data": [
        ["单元格1", "单元格2", "单元格3", "单元格4"],
        ["数据1", "数据2", "数据3", "数据4"]
      ]
    }
  ],
  "images": {
    "img_12345678": {
      "url": "images/img_12345678.png",
      "format": "png",
      "width": 800,
      "height": 600,
      "size": "150.23 KB",
      "context": "段落中的图片"
    }
  },
  "smartart": {
    "smartart_12345678": {
      "title": "组织架构图",
      "type": "hierarchy",
      "layout": "org_chart",
      "nodes": [
        {
          "text": "CEO",
          "level": 0,
          "children": ["node_001", "node_002"]
        }
      ]
    }
  },
  "embedded_objects": {
    "embedded_obj_12345678": {
      "type": "visio",
      "title": "流程图",
      "size": "2.5 MB",
      "preview_image": "images/embedded_preview_12345678.png"
    }
  },
  "processing_info": {
    "total_time": "2.34 seconds",
    "quick_mode": true,
    "warnings": [],
    "errors": []
  }
}
```

## 🔍 故障排除

### 常见错误解决方案

1. **ImportError: No module named 'docx'**
   ```bash
   pip install python-docx
   ```

2. **Wand相关错误**
   ```bash
   # macOS
   brew install imagemagick
   pip install Wand
   
   # Ubuntu
   sudo apt-get install imagemagick libmagickwand-dev
   pip install Wand
   ```

3. **内存不足错误**
   - 使用快速模式: `quick_mode=True`
   - 分批处理大文件
   - 增加系统内存

4. **文件权限错误**
   ```bash
   # 确保有读写权限
   chmod 755 Files/
   chmod 644 Files/*.docx
   ```

### 调试技巧

1. **启用详细日志**
   ```python
   import logging
   logging.basicConfig(level=logging.DEBUG)
   ```

2. **检查文件完整性**
   ```python
   from zipfile import ZipFile
   
   def check_docx_file(docx_path):
       try:
           with ZipFile(docx_path, 'r') as zip_file:
               # 检查必要文件
               required_files = ['word/document.xml', '[Content_Types].xml']
               for file in required_files:
                   if file not in zip_file.namelist():
                       print(f"❌ 缺少必要文件: {file}")
                       return False
               print("✅ DOCX文件结构正常")
               return True
       except Exception as e:
           print(f"❌ 文件损坏: {e}")
           return False
   ```

3. **性能分析**
   ```python
   import cProfile
   
   def profile_parsing(docx_path):
       """性能分析"""
       cProfile.run(
           f'parse_docx("{docx_path}", "output", quick_mode=True)',
           'profile_stats.txt'
       )
   ```

## 📚 API参考

### 主要函数

#### `parse_docx(docx_path, output_dir, quick_mode=True)`

解析单个DOCX文档。

**参数:**
- `docx_path` (str): DOCX文件路径
- `output_dir` (str): 输出目录路径  
- `quick_mode` (bool): 是否启用快速模式，默认True

**返回:**
- `dict` | `None`: 解析结果字典，失败时返回None

#### `process_docx_folder(input_folder, output_folder, quick_mode=True)`

批量处理文件夹中的DOCX文档。

**参数:**
- `input_folder` (str): 输入文件夹路径
- `output_folder` (str): 输出文件夹路径
- `quick_mode` (bool): 是否启用快速模式，默认True

**返回:**
- `int`: 成功处理的文件数量

---

更多详细信息请参考源代码注释和示例文件。
