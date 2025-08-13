#!/usr/bin/env python3
"""
快速解析示例
演示如何使用快速模式解析DOCX文档
"""

import os
import sys
import time

# 添加父目录到Python路径以便导入模块
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from src.parsers.document_parser import parse_docx

def quick_parse_example():
    """快速解析示例"""
    
    # 示例文件路径
    test_files = [
        "../Files/PLM2.0/BOM 审核申请.docx",
        "../Files/PLM2.0/CBB废弃申请.docx",
        "../Files/PLM2.0/CBB维护申请.docx",
        "../Files/PLM2.0/CBB新建申请V2.0.docx",
        "Files/PLM2.0/新ECR.docx", 
        "Files/PLM2.0/材料维护申请流程V2.0.docx"
    ]
    
    # 找一个存在的文件进行测试
    test_file = None
    for file_path in test_files:
        if os.path.exists(file_path):
            test_file = file_path
            break
    
    if not test_file:
        print("未找到测试文件，请确保有DOCX文件可供测试")
        return
    
    print(f"使用测试文件: {test_file}")
    
    # 创建输出目录
    output_dir = "quick_test_output"
    os.makedirs(output_dir, exist_ok=True)
    
    print("\n=== 快速模式解析 ===")
    start_time = time.time()
    
    # 使用快速模式解析（默认为True）
    result = parse_docx(test_file, output_dir, quick_mode=True)
    
    end_time = time.time()
    
    if result:
        print(f"✅ 解析成功！")
        print(f"⏱️  耗时: {end_time - start_time:.2f} 秒")
        
        # 显示解析结果统计
        images_count = len(result.get("images", {}))
        
        # 统计SmartArt和嵌入对象
        smartart_count = 0
        embedded_count = 0
        
        def count_content_types(content_list):
            sa_count = 0
            emb_count = 0
            for item in content_list:
                if isinstance(item, dict):
                    if item.get("type") == "smartart":
                        sa_count += 1
                    elif item.get("type") == "embedded_object":
                        emb_count += 1
                    elif item.get("type") == "section" and "content" in item:
                        child_sa, child_emb = count_content_types(item["content"])
                        sa_count += child_sa
                        emb_count += child_emb
                    elif "content" in item and isinstance(item["content"], list):
                        child_sa, child_emb = count_content_types(item["content"])
                        sa_count += child_sa
                        emb_count += child_emb
            return sa_count, emb_count
        
        for section in result.get("sections", []):
            section_sa, section_emb = count_content_types(section.get("content", []))
            smartart_count += section_sa
            embedded_count += section_emb
        
        print(f"📊 解析内容:")
        print(f"   - 图片: {images_count} 个")
        print(f"   - SmartArt: {smartart_count} 个")
        print(f"   - 嵌入对象: {embedded_count} 个")
        
        # 显示处理信息
        processing_info = result.get("processing_info", {})
        errors = processing_info.get("errors", [])
        warnings = processing_info.get("warnings", [])
        
        if errors:
            print(f"❌ 错误: {len(errors)} 个")
        if warnings:
            print(f"⚠️  警告: {len(warnings)} 个")
            
        print(f"📁 输出目录: {output_dir}")
        print(f"📄 详细结果保存在: {os.path.join(output_dir, 'document.json')}")
        
        # 显示快速模式的优势
        print(f"\n💡 快速模式优势:")
        print(f"   - 跳过耗时的EMF/WMF转换（节省最多20秒/图像）")
        print(f"   - 直接保存原格式文件，无转换失败风险")
        print(f"   - 适合快速提取文档结构和文本内容")
        print(f"   - 保留所有重要信息，仅图像格式保持原样")
        
    else:
        print(f"❌ 解析失败")

if __name__ == "__main__":
    quick_parse_example()
