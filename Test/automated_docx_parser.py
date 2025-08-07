#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
完全自动化的 DOCX 文档解析器
自动提取所有内容，包括文本、图片、表格、SmartArt 和嵌入对象（如 Visio 流程图）
并尝试将 EMF 格式自动转换为 PNG 格式
"""

import os
import sys
import json
import logging
import shutil
from pathlib import Path
from docx_parser import parse_docx, process_docx_folder
import emf_auto_converter

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def analyze_conversion_results(output_dir):
    """
    分析转换结果并生成报告
    """
    try:
        # 统计各种文件类型
        images_dir = os.path.join(output_dir, "images")
        if not os.path.exists(images_dir):
            return None
        
        stats = {
            "png_files": [],
            "emf_files": [],
            "jpg_files": [],
            "other_files": []
        }
        
        for filename in os.listdir(images_dir):
            file_path = os.path.join(images_dir, filename)
            if os.path.isfile(file_path):
                if filename.lower().endswith('.png'):
                    stats["png_files"].append(filename)
                elif filename.lower().endswith('.emf'):
                    stats["emf_files"].append(filename)
                elif filename.lower().endswith(('.jpg', '.jpeg')):
                    stats["jpg_files"].append(filename)
                else:
                    stats["other_files"].append(filename)
        
        return stats
        
    except Exception as e:
        logger.error(f"分析转换结果失败: {e}")
        return None

def create_user_guide(output_dir, stats):
    """
    为用户创建使用指南
    """
    try:
        guide_content = f"""# 文档解析结果使用指南

## 📊 文件统计

- **PNG 图片**: {len(stats['png_files'])} 个（可直接查看）
- **EMF 文件**: {len(stats['emf_files'])} 个（需要转换）
- **JPG 图片**: {len(stats['jpg_files'])} 个（可直接查看）
- **其他文件**: {len(stats['other_files'])} 个

## 📁 目录结构

```
{os.path.basename(output_dir)}/
├── document.json          # 完整的文档结构化数据
├── images/                 # 所有提取的图片
│   ├── *.png              # 可直接查看的图片
│   ├── *.emf              # 需要转换的 EMF 文件
│   └── conversion_info_*.json  # EMF 转换指导信息
├── smartart/              # SmartArt 图表数据
├── embedded_objects/      # 嵌入对象信息
└── emf_conversion_summary.json  # EMF 转换总结报告
```

## 🎯 如何查看 EMF 流程图

### 自动处理的部分
✅ 系统已自动提取了所有嵌入的 Visio 流程图
✅ 已尝试使用多种方法转换为 PNG 格式
✅ 生成了详细的转换指导信息

### 需要手动处理的 EMF 文件

对于以下 EMF 文件，您可以：

"""
        
        if stats['emf_files']:
            for emf_file in stats['emf_files']:
                guide_content += f"#### {emf_file}\n"
                guide_content += f"- **推荐方法**: 使用在线转换工具\n"
                guide_content += f"  1. 访问：https://convertio.co/emf-png/\n"
                guide_content += f"  2. 上传文件：`images/{emf_file}`\n"
                guide_content += f"  3. 下载 PNG 结果\n\n"
        else:
            guide_content += "🎉 所有文件都已成功转换为可查看格式！\n\n"
        
        guide_content += """
## 📖 查看文档内容

### 1. 结构化数据
- 打开 `document.json` 查看完整的文档结构
- 包含所有文本、表格、图片、SmartArt 和嵌入对象

### 2. 图片内容
- 浏览 `images/` 目录查看所有提取的图片
- PNG 和 JPG 文件可直接查看
- EMF 文件需要转换后查看

### 3. 嵌入对象
- 查看 `embedded_objects/` 目录了解嵌入对象详情
- 每个对象都有对应的 JSON 描述文件

## 🔧 进阶使用

### 编程访问
```python
import json

# 读取文档结构
with open('document.json', 'r', encoding='utf-8') as f:
    doc = json.load(f)

# 查找所有嵌入对象
embedded_objects = []
def find_embedded_objects(obj):
    if isinstance(obj, dict):
        if obj.get('type') == 'embedded_object':
            embedded_objects.append(obj)
        for value in obj.values():
            find_embedded_objects(value)
    elif isinstance(obj, list):
        for item in obj:
            find_embedded_objects(item)

find_embedded_objects(doc)
print(f"找到 {len(embedded_objects)} 个嵌入对象")
```

### 批量处理
- 使用 `python automated_docx_parser.py <输入目录>` 批量处理多个文档
- 结果会保存在各自的子目录中

## 📞 技术支持

如果遇到问题：
1. 检查 `emf_conversion_summary.json` 了解转换详情
2. 查看 `conversion_info_*.json` 文件获取特定文件的转换指导
3. 使用推荐的在线转换工具处理 EMF 文件

---
*由自动化 DOCX 解析器生成*
"""
        
        guide_path = os.path.join(output_dir, "使用指南.md")
        with open(guide_path, 'w', encoding='utf-8') as f:
            f.write(guide_content)
        
        logger.info(f"📖 使用指南已创建: {guide_path}")
        return guide_path
        
    except Exception as e:
        logger.error(f"创建使用指南失败: {e}")
        return None

def process_single_file(docx_path, output_base_dir=None):
    """
    处理单个 DOCX 文件的完整流水线
    """
    logger.info(f"🚀 开始处理单个文件: {docx_path}")
    
    # 确定输出目录
    if output_base_dir is None:
        output_base_dir = "Files/automated_output"
    
    filename = Path(docx_path).stem
    safe_name = filename.replace(' ', '_').replace('.', '_')
    output_dir = os.path.join(output_base_dir, safe_name)
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    logger.info(f"📁 输出目录: {output_dir}")
    
    # 第一步：解析文档
    logger.info("🔍 第一步：解析 DOCX 文档...")
    result = parse_docx(docx_path, output_dir)
    
    if not result:
        logger.error("❌ 文档解析失败")
        return None
    
    # 保存 JSON 结果
    json_path = os.path.join(output_dir, "document.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    logger.info("✅ 文档解析完成")
    
    # 第二步：EMF 自动转换
    logger.info("🔄 第二步：尝试自动转换 EMF 文件...")
    emf_auto_converter.main(output_dir)
    
    # 第三步：分析结果
    logger.info("📊 第三步：分析转换结果...")
    stats = analyze_conversion_results(output_dir)
    
    # 第四步：创建用户指南
    logger.info("📖 第四步：生成使用指南...")
    if stats:
        create_user_guide(output_dir, stats)
    
    # 生成最终报告
    logger.info("\n" + "="*60)
    logger.info("🎉 处理完成！")
    logger.info("="*60)
    
    if stats:
        logger.info(f"📊 文件统计:")
        logger.info(f"   PNG 图片: {len(stats['png_files'])} 个")
        logger.info(f"   EMF 文件: {len(stats['emf_files'])} 个")
        logger.info(f"   JPG 图片: {len(stats['jpg_files'])} 个")
        logger.info(f"   其他文件: {len(stats['other_files'])} 个")
        
        if stats['emf_files']:
            logger.info(f"\n⚠️  需要手动转换的 EMF 文件:")
            for emf_file in stats['emf_files']:
                logger.info(f"   - {emf_file}")
            logger.info(f"\n💡 建议使用在线工具: https://convertio.co/emf-png/")
        else:
            logger.info(f"\n🎉 所有图片都已成功转换为可查看格式！")
    
    logger.info(f"\n📁 结果保存在: {output_dir}")
    logger.info(f"📖 使用指南: {os.path.join(output_dir, '使用指南.md')}")
    
    return output_dir

def process_folder(input_folder, output_base_dir=None):
    """
    批量处理文件夹中的所有 DOCX 文件
    """
    logger.info(f"🚀 开始批量处理文件夹: {input_folder}")
    
    if output_base_dir is None:
        output_base_dir = "Files/automated_batch_output"
    
    # 第一步：批量解析所有文档
    logger.info("🔍 第一步：批量解析所有 DOCX 文档...")
    processed_count = process_docx_folder(input_folder, output_base_dir)
    
    if processed_count == 0:
        logger.error("❌ 没有成功处理任何文档")
        return None
    
    # 第二步：对每个输出目录运行 EMF 转换
    logger.info("🔄 第二步：对所有文档尝试 EMF 转换...")
    
    for item in os.listdir(output_base_dir):
        item_path = os.path.join(output_base_dir, item)
        if os.path.isdir(item_path) and item != "summary.json":
            logger.info(f"   处理目录: {item}")
            emf_auto_converter.main(item_path)
            
            # 为每个目录生成使用指南
            stats = analyze_conversion_results(item_path)
            if stats:
                create_user_guide(item_path, stats)
    
    logger.info("\n" + "="*60)
    logger.info("🎉 批量处理完成！")
    logger.info("="*60)
    logger.info(f"📁 结果保存在: {output_base_dir}")
    logger.info(f"📋 总体报告: {os.path.join(output_base_dir, 'summary.json')}")
    
    return output_base_dir

def main():
    """
    主函数：根据参数决定处理单个文件还是文件夹
    """
    print("🤖 自动化 DOCX 文档解析器")
    print("=" * 50)
    
    if len(sys.argv) < 2:
        print("使用方法:")
        print("  处理单个文件: python automated_docx_parser.py <docx_file>")
        print("  处理文件夹:   python automated_docx_parser.py <folder_path>")
        print("\n示例:")
        print("  python automated_docx_parser.py 'Files/PLM2.0/BOM 审核申请.docx'")
        print("  python automated_docx_parser.py Files/PLM2.0")
        sys.exit(1)
    
    input_path = sys.argv[1]
    
    if not os.path.exists(input_path):
        logger.error(f"❌ 路径不存在: {input_path}")
        sys.exit(1)
    
    try:
        if os.path.isfile(input_path) and input_path.lower().endswith('.docx'):
            # 处理单个文件
            result = process_single_file(input_path)
        elif os.path.isdir(input_path):
            # 处理文件夹
            result = process_folder(input_path)
        else:
            logger.error(f"❌ 不支持的文件类型或路径: {input_path}")
            sys.exit(1)
        
        if result:
            print(f"\n🎊 成功完成！结果保存在: {result}")
        else:
            print(f"\n❌ 处理失败")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print(f"\n⏹️  用户取消操作")
        sys.exit(0)
    except Exception as e:
        logger.error(f"❌ 处理过程中发生错误: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
