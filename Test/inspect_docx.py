#!/usr/bin/env python3
"""
检查DOCX文件结构
"""

import zipfile
import os

def inspect_docx_structure(docx_path):
    """检查DOCX文件的内部结构"""
    
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # 列出所有文件
        all_files = docx_zip.namelist()
        print(f"DOCX文件包含 {len(all_files)} 个文件:")
        
        for file in sorted(all_files):
            print(f"  {file}")
        
        # 检查是否有diagrams目录
        diagram_files = [f for f in all_files if 'diagram' in f.lower()]
        print(f"\n包含diagram的文件: {len(diagram_files)}")
        for file in diagram_files:
            print(f"  {file}")
        
        # 检查关系文件
        rels_files = [f for f in all_files if 'rels' in f]
        print(f"\n关系文件: {len(rels_files)}")
        for file in rels_files:
            print(f"  {file}")
            
            try:
                content = docx_zip.read(file)
                print(f"    文件大小: {len(content)} 字节")
                if len(content) < 2000:  # 如果文件不太大，显示内容
                    print(f"    内容:\n{content.decode('utf-8')}")
            except Exception as e:
                print(f"    读取失败: {e}")

if __name__ == "__main__":
    docx_path = "Files/PLM2.0/BOM 审核申请.docx"
    
    if os.path.exists(docx_path):
        print(f"检查文档: {docx_path}")
        inspect_docx_structure(docx_path)
    else:
        print(f"文档不存在: {docx_path}")
