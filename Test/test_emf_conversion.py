#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试自动化 EMF 到 PNG 转换功能
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from docx_parser import parse_docx
import json

def test_single_docx(docx_path, output_dir):
    """测试单个 DOCX 文件的解析和转换"""
    print(f"🔍 开始解析文档：{docx_path}")
    print(f"📁 输出目录：{output_dir}")
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 解析文档
    result = parse_docx(docx_path, output_dir)
    
    if result:
        print("✅ 文档解析成功！")
        
        # 保存解析结果
        result_file = os.path.join(output_dir, "document.json")
        with open(result_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"📄 解析结果已保存到：{result_file}")
        
        # 查找嵌入对象
        embedded_objects = []
        def find_embedded_objects(content):
            if isinstance(content, list):
                for item in content:
                    find_embedded_objects(item)
            elif isinstance(content, dict):
                if content.get('type') == 'embedded_object':
                    embedded_objects.append(content)
                for value in content.values():
                    find_embedded_objects(value)
        
        find_embedded_objects(result)
        
        if embedded_objects:
            print(f"\n🎯 发现 {len(embedded_objects)} 个嵌入对象：")
            for obj in embedded_objects:
                print(f"   - {obj.get('description', 'Unknown')} ({obj.get('width', 'Unknown')} x {obj.get('height', 'Unknown')})")
                if 'preview_image' in obj:
                    image_path = os.path.join(output_dir, obj['preview_image'])
                    if os.path.exists(image_path):
                        if image_path.endswith('.png'):
                            print(f"     ✅ PNG 文件已生成：{obj['preview_image']}")
                        else:
                            print(f"     ⚠️  原始格式保存：{obj['preview_image']}")
                    else:
                        print(f"     ❌ 文件缺失：{obj['preview_image']}")
        else:
            print("\n📝 未发现嵌入对象")
            
        # 检查图像目录
        images_dir = os.path.join(output_dir, "images")
        if os.path.exists(images_dir):
            image_files = os.listdir(images_dir)
            png_files = [f for f in image_files if f.endswith('.png')]
            emf_files = [f for f in image_files if f.endswith('.emf')]
            
            print(f"\n📊 图像文件统计：")
            print(f"   - PNG 文件：{len(png_files)} 个")
            print(f"   - EMF 文件：{len(emf_files)} 个")
            print(f"   - 其他文件：{len(image_files) - len(png_files) - len(emf_files)} 个")
            
            if png_files:
                print(f"\n🎉 成功转换的 PNG 文件：")
                for png_file in png_files:
                    print(f"   - {png_file}")
            
            if emf_files:
                print(f"\n⚠️  需要进一步处理的 EMF 文件：")
                for emf_file in emf_files:
                    print(f"   - {emf_file}")
        
        return True
    else:
        print("❌ 文档解析失败")
        return False

def main():
    """主函数"""
    print("🧪 === EMF 自动转换测试 === 🧪\n")
    
    # 测试文件
    docx_file = "Files/PLM2.0/BOM 审核申请.docx"
    output_dir = "Files/test_automated_conversion"
    
    if not os.path.exists(docx_file):
        print(f"❌ 文件不存在：{docx_file}")
        return
    
    # 运行测试
    success = test_single_docx(docx_file, output_dir)
    
    if success:
        print(f"\n🎊 测试完成！请检查输出目录：{output_dir}")
    else:
        print(f"\n💥 测试失败！")

if __name__ == "__main__":
    main()
