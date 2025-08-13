#!/usr/bin/env python3
"""
完整的文档处理演示
展示从DOCX解析到标准化文本输出的完整流程
"""

import sys
import os
import json
import logging
import time
from pathlib import Path

# 添加src目录到Python路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.parsers.document_parser import parse_docx
from src.processors.text_processor import process_document_to_text
from src.utils.text_utils import safe_filename

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def print_separator(title, char="=", width=80):
    """打印分隔线"""
    print(f"\n{char * width}")
    print(f"{title:^{width}}")
    print(f"{char * width}")

def process_demo_file(docx_path, output_base_dir="demo_output"):
    """
    处理演示文件，展示完整的处理流程
    """
    try:
        # 创建输出目录
        file_basename = os.path.splitext(os.path.basename(docx_path))[0]
        safe_name = safe_filename(file_basename)
        output_dir = os.path.join(output_base_dir, safe_name)
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"📁 输入文件: {docx_path}")
        print(f"📁 输出目录: {output_dir}")
        
        # 步骤1：解析DOCX文档
        print_separator("步骤1: 解析DOCX文档", "-", 50)
        start_time = time.time()
        
        document_structure = parse_docx(docx_path, output_dir, quick_mode=True)
        
        parse_time = time.time() - start_time
        
        if not document_structure:
            print("❌ 文档解析失败")
            return False
        
        print(f"✅ 文档解析成功，耗时: {parse_time:.2f}秒")
        
        # 保存JSON结果
        json_path = os.path.join(output_dir, "document.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(document_structure, f, ensure_ascii=False, indent=2)
        print(f"📄 JSON结构已保存到: {json_path}")
        
        # 显示解析统计
        sections = document_structure.get("sections", [])
        images = document_structure.get("images", {})
        processing_info = document_structure.get("processing_info", {})
        
        print(f"📊 解析统计:")
        print(f"   📝 章节数量: {len(sections)}")
        print(f"   🖼️  图片数量: {len(images)}")
        print(f"   ⚠️  警告数量: {len(processing_info.get('warnings', []))}")
        print(f"   ❌ 错误数量: {len(processing_info.get('errors', []))}")
        
        # 步骤2：处理为标准化文本
        print_separator("步骤2: 生成标准化文本", "-", 50)
        start_time = time.time()
        
        processed_text = process_document_to_text(document_structure, safe_name)
        
        process_time = time.time() - start_time
        
        if not processed_text:
            print("❌ 文本处理失败")
            return False
        
        print(f"✅ 文本处理成功，耗时: {process_time:.2f}秒")
        
        # 保存处理后的文本
        text_path = os.path.join(output_dir, "processed_text.txt")
        with open(text_path, "w", encoding="utf-8") as f:
            f.write(processed_text)
        print(f"📄 标准化文本已保存到: {text_path}")
        
        # 显示文本统计
        lines = processed_text.split('\n')
        non_empty_lines = [line for line in lines if line.strip()]
        
        print(f"📊 文本统计:")
        print(f"   📏 总字符数: {len(processed_text)}")
        print(f"   📄 总行数: {len(lines)}")
        print(f"   📝 非空行数: {len(non_empty_lines)}")
        
        # 步骤3：展示处理结果
        print_separator("步骤3: 处理结果预览", "-", 50)
        
        # 显示文本预览
        preview_length = 800
        if len(processed_text) > preview_length:
            preview_text = processed_text[:preview_length] + "\n... [省略剩余内容]"
        else:
            preview_text = processed_text
        
        print("📖 标准化文本预览:")
        print("─" * 60)
        print(preview_text)
        print("─" * 60)
        
        # 步骤4：总结
        print_separator("处理完成", "=", 50)
        total_time = parse_time + process_time
        print(f"⏱️  总处理时间: {total_time:.2f}秒")
        print(f"📁 所有结果保存在: {output_dir}")
        print(f"📄 JSON结构文件: document.json")
        print(f"📄 标准化文本文件: processed_text.txt")
        
        return True
        
    except Exception as e:
        print(f"❌ 处理失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主函数"""
    print_separator("DOCX文档处理演示", "=", 80)
    print("🎯 功能：将DOCX文档解析并转换为标准化文本格式")
    print("📋 基于：需求规格说明书3.1-3.10节的处理规则")
    
    # 演示文件列表
    demo_files = [
        "Files/examples/demo.docx",
        "Files/PLM2.0/BOM 审核申请.docx"
    ]
    
    # 创建演示输出目录
    output_base_dir = "demo_output"
    os.makedirs(output_base_dir, exist_ok=True)
    
    success_count = 0
    total_count = 0
    
    for demo_file in demo_files:
        if os.path.exists(demo_file):
            print_separator(f"处理文件: {os.path.basename(demo_file)}", "🔄", 80)
            total_count += 1
            
            if process_demo_file(demo_file, output_base_dir):
                success_count += 1
                print("✅ 处理成功！")
            else:
                print("❌ 处理失败！")
        else:
            print(f"⚠️  演示文件不存在: {demo_file}")
    
    # 最终总结
    print_separator("演示总结", "=", 80)
    print(f"📊 处理结果: {success_count}/{total_count} 个文件成功")
    
    if success_count > 0:
        print(f"📁 演示结果保存在: {output_base_dir}/")
        print("📖 查看各文件夹中的 processed_text.txt 了解处理结果")
        
        # 显示功能特性
        print("\n🎯 处理器功能特性:")
        print("   ✅ 自动识别章节结构")
        print("   ✅ 标准化标签格式 (<|TAG|><|/TAG|>)")
        print("   ✅ 表格数据结构化处理")
        print("   ✅ 图片Markdown格式转换")
        print("   ✅ 符合需求规格说明书规范")
        
    print("\n🎉 演示完成！")

if __name__ == "__main__":
    main()
