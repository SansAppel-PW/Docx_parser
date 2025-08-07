#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试单个DOCX文件的解析，专门用于验证嵌入对象解析
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from docx_parser import parse_docx
import json

def test_single_file():
    docx_path = "Files/PLM2.0/BOM 审核申请.docx"
    output_dir = "Files/test_embedded_objects"
    
    print(f"测试文件: {docx_path}")
    print(f"输出目录: {output_dir}")
    
    if not os.path.exists(docx_path):
        print(f"文件不存在: {docx_path}")
        return
    
    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)
    
    # 解析文档
    result = parse_docx(docx_path, output_dir)
    
    if result:
        # 保存结果
        output_file = os.path.join(output_dir, "document.json")
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"解析成功，结果保存到: {output_file}")
        
        # 统计各种内容类型
        sections = result.get("sections", [])
        
        def count_content_types(sections):
            counts = {
                "paragraphs": 0,
                "tables": 0,
                "images": 0,
                "smartart": 0,
                "embedded_objects": 0,
                "list_items": 0
            }
            
            def count_in_content(content_list):
                for item in content_list:
                    if item.get("type") == "paragraph":
                        counts["paragraphs"] += 1
                    elif item.get("type") == "table":
                        counts["tables"] += 1
                    elif item.get("type") == "image":
                        counts["images"] += 1
                    elif item.get("type") == "smartart":
                        counts["smartart"] += 1
                    elif item.get("type") == "embedded_object":
                        counts["embedded_objects"] += 1
                    elif item.get("type") == "list_item":
                        counts["list_items"] += 1
                    elif item.get("type") == "section":
                        count_in_content(item.get("content", []))
            
            for section in sections:
                count_in_content(section.get("content", []))
            
            return counts
        
        counts = count_content_types(sections)
        
        print("\n=== 内容统计 ===")
        for content_type, count in counts.items():
            print(f"{content_type}: {count}")
        
        # 专门查找嵌入对象
        print("\n=== 嵌入对象详情 ===")
        embedded_found = False
        
        def find_embedded_objects(content_list, path=""):
            nonlocal embedded_found
            for i, item in enumerate(content_list):
                if item.get("type") == "embedded_object":
                    embedded_found = True
                    print(f"位置: {path}[{i}]")
                    print(f"  类型: {item.get('object_type', 'unknown')}")
                    print(f"  描述: {item.get('description', 'N/A')}")
                    print(f"  程序ID: {item.get('prog_id', 'N/A')}")
                    print(f"  尺寸: {item.get('width', 'N/A')} x {item.get('height', 'N/A')}")
                    print(f"  上下文: {item.get('context', 'N/A')}")
                    if 'file_path' in item:
                        print(f"  详细信息文件: {item['file_path']}")
                elif item.get("type") == "section":
                    find_embedded_objects(item.get("content", []), f"{path}/section:{item.get('title', 'unknown')}")
        
        for i, section in enumerate(sections):
            find_embedded_objects(section.get("content", []), f"section[{i}]:{section.get('title', 'unknown')}")
        
        if not embedded_found:
            print("未找到嵌入对象")
        
    else:
        print("解析失败")

if __name__ == "__main__":
    test_single_file()
