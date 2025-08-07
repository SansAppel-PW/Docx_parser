#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试SmartArt解析功能
"""

import os
import sys
import json
from docx_parser import parse_docx

def test_smartart_parsing():
    """测试SmartArt解析功能"""
    
    # 测试包含SmartArt的文档
    test_file = "Files/PLM2.0/BOM 审核申请.docx"
    output_dir = "test_smartart_output"
    
    if not os.path.exists(test_file):
        print(f"测试文件不存在: {test_file}")
        return False
    
    print(f"开始解析文档: {test_file}")
    
    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)
    
    # 解析文档
    try:
        result = parse_docx(test_file, output_dir)
        
        if not result:
            print("解析失败")
            return False
        
        # 保存解析结果
        result_file = os.path.join(output_dir, "document_with_smartart.json")
        with open(result_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"解析完成，结果保存到: {result_file}")
        
        # 统计SmartArt数量
        smartart_count = count_smartart_elements(result)
        print(f"发现 {smartart_count} 个SmartArt元素")
        
        # 显示SmartArt详细信息
        show_smartart_details(result)
        
        return True
        
    except Exception as e:
        print(f"解析过程中出错: {e}")
        import traceback
        traceback.print_exc()
        return False

def count_smartart_elements(data):
    """递归统计SmartArt元素数量"""
    count = 0
    
    if isinstance(data, dict):
        if data.get("type") == "smartart":
            count += 1
        for value in data.values():
            count += count_smartart_elements(value)
    elif isinstance(data, list):
        for item in data:
            count += count_smartart_elements(item)
    
    return count

def show_smartart_details(data):
    """显示SmartArt详细信息"""
    smartarts = find_smartart_elements(data)
    
    if not smartarts:
        print("未发现SmartArt元素")
        return
    
    print(f"\nSmartArt详细信息:")
    print("=" * 50)
    
    for i, smartart in enumerate(smartarts, 1):
        print(f"\nSmartArt {i}:")
        print(f"  ID: {smartart.get('id', 'N/A')}")
        print(f"  类型: {smartart.get('diagram_type', 'unknown')}")
        print(f"  上下文: {smartart.get('context', 'N/A')}")
        
        text_content = smartart.get('text_content', [])
        if text_content:
            print(f"  文本内容 ({len(text_content)} 项):")
            for j, text_node in enumerate(text_content):
                print(f"    {j+1}. {text_node.get('text', '')}")
        else:
            print("  无文本内容")
        
        file_path = smartart.get('file_path', '')
        if file_path:
            print(f"  详细数据文件: {file_path}")

def find_smartart_elements(data):
    """递归查找所有SmartArt元素"""
    smartarts = []
    
    if isinstance(data, dict):
        if data.get("type") == "smartart":
            smartarts.append(data)
        for value in data.values():
            smartarts.extend(find_smartart_elements(value))
    elif isinstance(data, list):
        for item in data:
            smartarts.extend(find_smartart_elements(item))
    
    return smartarts

if __name__ == "__main__":
    print("SmartArt解析测试")
    print("=" * 50)
    
    success = test_smartart_parsing()
    
    if success:
        print("\n✅ 测试成功完成")
    else:
        print("\n❌ 测试失败")
        sys.exit(1)
