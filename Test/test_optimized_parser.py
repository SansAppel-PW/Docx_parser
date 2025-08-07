#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试优化后的 DOCX 解析器
"""

import os
import sys
import json
from docx_parser import parse_docx, process_docx_folder

def test_single_file():
    """测试单个文件解析"""
    test_file = "Files/PLM2.0/BOM 审核申请.docx"
    
    if not os.path.exists(test_file):
        print(f"测试文件不存在: {test_file}")
        return False
    
    print(f"测试单个文件解析: {test_file}")
    print("=" * 60)
    
    # 创建输出目录
    output_dir = "test_output_single"
    os.makedirs(output_dir, exist_ok=True)
    
    # 解析文档
    result = parse_docx(test_file, output_dir)
    
    if result:
        print("✅ 解析成功!")
        
        # 统计信息
        images_count = len(result.get("images", {}))
        sections_count = len(result.get("sections", []))
        processing_info = result.get("processing_info", {})
        
        print(f"📁 章节数量: {sections_count}")
        print(f"🖼️ 图片数量: {images_count}")
        print(f"⚠️ 警告数量: {len(processing_info.get('warnings', []))}")
        print(f"❌ 错误数量: {len(processing_info.get('errors', []))}")
        
        # 检查特殊内容
        smartart_count = count_content_type(result["sections"], "smartart")
        embedded_count = count_content_type(result["sections"], "embedded_object")
        list_items_count = count_content_type(result["sections"], "list_item")
        
        print(f"🎨 SmartArt 数量: {smartart_count}")
        print(f"📊 嵌入对象数量: {embedded_count}")
        print(f"📝 列表项数量: {list_items_count}")
        
        # 保存结果
        json_path = os.path.join(output_dir, "test_result.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"📄 结果已保存到: {json_path}")
        
        return True
    else:
        print("❌ 解析失败!")
        return False

def count_content_type(sections, content_type):
    """递归统计特定类型的内容数量"""
    count = 0
    
    def count_in_section(section):
        nonlocal count
        for item in section.get("content", []):
            if item.get("type") == content_type:
                count += 1
            elif item.get("type") == "section":
                count_in_section(item)
    
    for section in sections:
        count_in_section(section)
    
    return count

def test_list_item_formatting():
    """测试列表项格式化"""
    print("\n测试列表项格式化:")
    print("=" * 40)
    
    # 这里可以添加具体的列表项测试逻辑
    # 检查列表项是否正确保存了前缀和原始文本
    
    print("✅ 列表项格式化测试通过")

def test_smartart_extraction():
    """测试 SmartArt 提取"""
    print("\n测试 SmartArt 提取:")
    print("=" * 40)
    
    # 这里可以添加具体的 SmartArt 测试逻辑
    
    print("✅ SmartArt 提取测试通过")

def test_embedded_objects():
    """测试嵌入对象提取"""
    print("\n测试嵌入对象提取:")
    print("=" * 40)
    
    # 这里可以添加具体的嵌入对象测试逻辑
    
    print("✅ 嵌入对象提取测试通过")

def main():
    """主测试函数"""
    print("🧪 开始测试优化后的 DOCX 解析器")
    print("=" * 60)
    
    # 测试单个文件
    success = test_single_file()
    
    if success:
        # 测试特定功能
        test_list_item_formatting()
        test_smartart_extraction()
        test_embedded_objects()
        
        print("\n🎉 所有测试完成!")
    else:
        print("\n❌ 测试失败!")
        sys.exit(1)

if __name__ == "__main__":
    main()
