#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
使用改进后的SmartArt解析功能重新处理所有文档
"""

import os
import sys
from docx_parser import process_docx_folder

def main():
    """
    使用改进后的解析器重新处理所有文档
    """
    print("SmartArt解析功能增强版 - DOCX文档解析器")
    print("=" * 60)
    print("新增功能:")
    print("✅ 支持SmartArt图表检测和解析")
    print("✅ 提取SmartArt中的文本内容")
    print("✅ 识别SmartArt图表类型")
    print("✅ 保存SmartArt详细数据到独立文件")
    print("=" * 60)
    
    # 输入和输出目录
    input_folder = "Files/PLM2.0"
    output_base_dir = "Files/structured_docs_with_smartart"
    
    # 检查输入目录
    if not os.path.exists(input_folder):
        print(f"❌ 输入目录不存在: {input_folder}")
        sys.exit(1)
    
    print(f"📁 输入目录: {input_folder}")
    print(f"📁 输出目录: {output_base_dir}")
    print()
    
    # 处理所有文档
    try:
        processed_count = process_docx_folder(input_folder, output_base_dir)
        
        if processed_count > 0:
            print(f"\n🎉 处理完成！成功处理了 {processed_count} 个文档")
            print(f"📊 结果保存在: {output_base_dir}")
            
            # 统计SmartArt数量
            total_smartart = count_smartart_files(output_base_dir)
            print(f"📈 总共发现 {total_smartart} 个SmartArt图表")
            
        else:
            print("❌ 没有成功处理任何文档")
            
    except Exception as e:
        print(f"❌ 处理过程中出错: {e}")
        sys.exit(1)

def count_smartart_files(output_dir):
    """
    统计生成的SmartArt文件数量
    """
    count = 0
    try:
        for root, dirs, files in os.walk(output_dir):
            if 'smartart' in root:
                count += len([f for f in files if f.endswith('.json')])
    except:
        pass
    return count

if __name__ == "__main__":
    main()
