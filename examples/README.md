# 示例和演示

本目录包含DOCX解析器的完整使用示例，帮助您快速理解和使用项目的各种功能。

## 📋 示例文件概览

### 🚀 基础演示 - `demo.py`
**功能**: 展示项目的核心功能和基础用法
```bash
python examples/demo.py
```

**演示内容**:
- 单文档解析流程
- JSON输出结果查看
- 基础错误处理
- 处理时间统计

**适用场景**: 初次使用、功能验证、快速测试

---

### 📝 完整文本处理演示 - `demo_text_processing.py`
**功能**: 完整的文档处理到标准化文本输出的流程
```bash
python examples/demo_text_processing.py
```

**演示内容**:
- 完整的解析到文本转换流程
- 标准化文本格式展示
- 图片和嵌入对象处理
- 详细的处理步骤说明

**适用场景**: 了解完整处理流程、学习标准化输出格式

---

### 🧪 功能测试 - `test_text_processing.py`
**功能**: 文本处理功能的综合测试和验证
```bash
python examples/test_text_processing.py
```

**测试内容**:
- 文本处理器各项功能测试
- 输出格式验证
- 边界情况处理
- 性能基准测试

**适用场景**: 功能验证、回归测试、性能评估

---

## 🎯 使用建议

### 新用户学习路径
1. **第一步**: 运行 `demo.py` 了解基础功能
2. **第二步**: 运行 `demo_text_processing.py` 学习完整流程
3. **第三步**: 查看生成的输出文件理解格式
4. **第四步**: 运行 `test_text_processing.py` 验证功能

### 开发者参考路径
1. **代码参考**: 查看示例代码了解API使用方法
2. **功能测试**: 使用测试脚本验证修改后的功能
3. **自定义开发**: 基于示例代码开发自定义功能
4. **性能优化**: 参考性能测试进行优化

---

## 📁 示例输出结构

运行示例后，会在以下位置生成输出文件：

```
demo_output/
├── examples_parsed/           # demo.py输出
│   ├── demo/
│   │   ├── document.json     # JSON结构化数据
│   │   ├── processed_text.txt # 标准化文本
│   │   └── images/           # 提取的图片
│   └── summary.json          # 处理统计
│
├── text_demo_output/         # demo_text_processing.py输出
│   └── [类似结构]
│
└── test_output/              # test_text_processing.py输出
    └── [测试结果文件]
```

---

## 🔧 自定义示例

### 创建自己的示例
```python
#!/usr/bin/env python3
"""
自定义示例模板
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from src.parsers.document_parser import parse_docx
from src.processors.text_processor import process_document_to_text

def my_custom_demo():
    """自定义演示函数"""
    print("🚀 我的自定义演示")
    
    # 输入文件路径
    input_file = "../Files/examples/demo.docx"
    output_dir = "my_output"
    
    # 解析文档
    result = parse_docx(input_file, output_dir, quick_mode=True)
    
    if result:
        # 生成文本
        text_file = process_document_to_text(result, output_dir + "/demo")
        print(f"✅ 处理完成: {text_file}")
    else:
        print("❌ 处理失败")

if __name__ == "__main__":
    my_custom_demo()
```

### 批量处理示例
```python
from src.parsers.batch_processor import process_docx_folder

def batch_demo():
    """批量处理演示"""
    input_folder = "../Files/PLM2.0"
    output_folder = "batch_output"
    
    count = process_docx_folder(
        input_folder, 
        output_folder, 
        quick_mode=True
    )
    
    print(f"批量处理完成，处理了 {count} 个文件")
```

---

## 📊 示例执行结果

### 预期输出示例
```
🚀 DOCX解析器演示 - 基础功能
====================================

📂 输入文件: Files/examples/demo.docx
📁 输出目录: demo_output/examples_parsed/demo

⏱️  开始解析...
✅ 文档解析完成
📊 处理统计:
   - 段落数: 15
   - 表格数: 2  
   - 图片数: 3
   - 处理时间: 0.5秒

📄 生成文件:
   - JSON数据: demo_output/examples_parsed/demo/document.json
   - 标准文本: demo_output/examples_parsed/demo/processed_text.txt
   - 图片文件: demo_output/examples_parsed/demo/images/

🎉 演示完成！
```

---

## ⚠️ 注意事项

### 运行环境
- 确保在项目根目录运行示例
- 安装了所有必需的依赖包
- 有足够的磁盘空间存储输出文件

### 示例文件
- 示例使用 `Files/examples/demo.docx` 作为测试文件
- 如果文件不存在，请确保示例文档完整
- 可以替换为自己的DOCX文件进行测试

### 输出清理
```bash
# 清理示例输出（如果需要）
rm -rf demo_output/
rm -rf text_demo_output/
rm -rf test_output/
```

---

## 🔍 故障排除

### 常见问题

**Q: 运行示例时提示模块找不到**
```bash
# 解决方案：确保在项目根目录运行
cd /path/to/Docx_parser
python examples/demo.py
```

**Q: 找不到示例文件**
```bash
# 检查文件是否存在
ls Files/examples/demo.docx

# 如果不存在，可以使用其他docx文件替代
```

**Q: 输出目录权限错误**
```bash
# 确保有写入权限
chmod 755 .
mkdir -p demo_output
```

**Q: 处理速度很慢**
```python
# 启用快速模式
result = parse_docx(file, output, quick_mode=True)
```

---

## 📚 进一步学习

- 📖 查看 [使用指南](../docs/usage-guide.md) 了解详细API
- 🔧 参考 [文本处理指南](../docs/text-processor-guide.md) 了解内部机制
- 🧪 查看 [测试目录](../tests/) 了解更多测试用例
- 🏗️ 参考 [需求文档](../REQUIREMENTS.md) 了解技术规格

---

**祝您使用愉快！如有问题请查看文档或提交Issue。** 🚀
