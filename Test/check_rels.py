#!/usr/bin/env python3
"""
检查文档关系文件内容
"""

import zipfile
import os
from lxml import etree

def check_document_rels(docx_path):
    """检查主文档关系文件"""
    
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # 读取主文档关系文件
        rels_content = docx_zip.read('word/_rels/document.xml.rels')
        print(f"关系文件内容 ({len(rels_content)} 字节):")
        print(rels_content.decode('utf-8'))
        
        # 解析XML
        rels_tree = etree.fromstring(rels_content)
        
        # 命名空间
        namespaces = {
            'r': 'http://schemas.openxmlformats.org/package/2006/relationships'
        }
        
        # 查找所有关系
        all_rels = rels_tree.xpath('.//r:Relationship', namespaces=namespaces)
        print(f"\n解析出 {len(all_rels)} 个关系:")
        
        for rel in all_rels:
            rel_type = rel.get('Type', '')
            rel_id = rel.get('Id', '')
            target = rel.get('Target', '')
            print(f"  {rel_id}: {target}")
            print(f"    类型: {rel_type}")
            
            if 'diagram' in target.lower() or 'diagram' in rel_type.lower():
                print(f"    *** 这是一个diagram相关的关系 ***")

if __name__ == "__main__":
    docx_path = "Files/PLM2.0/BOM 审核申请.docx"
    
    if os.path.exists(docx_path):
        print(f"检查文档: {docx_path}")
        check_document_rels(docx_path)
    else:
        print(f"文档不存在: {docx_path}")
