#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
专门分析DOCX中的object元素和嵌入对象
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from lxml import etree
import json

def analyze_object_element(docx_path):
    """分析DOCX文件中的object元素"""
    print(f"分析文件中的object元素: {docx_path}")
    
    with zipfile.ZipFile(docx_path, 'r') as zip_file:
        document_xml = zip_file.read('word/document.xml')
        root = etree.fromstring(document_xml)
        
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'o': 'urn:schemas-microsoft-com:office:office',
            'v': 'urn:schemas-microsoft-com:vml',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
        
        # 查找所有object元素
        objects = root.findall('.//w:object', namespaces)
        print(f"找到 {len(objects)} 个object元素")
        
        for i, obj in enumerate(objects):
            print(f"\n=== Object {i+1} ===")
            
            # 打印完整的XML
            xml_str = etree.tostring(obj, encoding='unicode', pretty_print=True)
            print("完整XML内容:")
            print(xml_str)
            
            # 分析object的属性和子元素
            print("\nObject分析:")
            print(f"  Tag: {obj.tag}")
            print(f"  Attributes: {obj.attrib}")
            
            # 查找嵌入的文件引用
            ole_objects = obj.findall('.//o:OLEObject', namespaces)
            for ole in ole_objects:
                print(f"  OLEObject attributes: {ole.attrib}")
                
                # 获取关系ID
                r_id = ole.get(f'{{{namespaces["r"]}}}id')
                if r_id:
                    print(f"  关系ID: {r_id}")
                    
                    # 检查关系文件
                    analyze_relationship(zip_file, r_id)
            
            # 查找shape元素
            shapes = obj.findall('.//v:shape', namespaces)
            for shape in shapes:
                print(f"  Shape attributes: {shape.attrib}")
                
                # 查找imagedata
                imagedatas = shape.findall('.//v:imagedata', namespaces)
                for imagedata in imagedatas:
                    print(f"  ImageData attributes: {imagedata.attrib}")

def analyze_relationship(zip_file, r_id):
    """分析关系文件"""
    try:
        # 读取关系文件
        rels_xml = zip_file.read('word/_rels/document.xml.rels')
        rels_root = etree.fromstring(rels_xml)
        
        namespaces = {
            'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
        }
        
        # 查找指定的关系
        relationships = rels_root.findall('.//rel:Relationship', namespaces)
        for rel in relationships:
            if rel.get('Id') == r_id:
                print(f"  关系类型: {rel.get('Type')}")
                print(f"  目标文件: {rel.get('Target')}")
                
                # 如果目标文件存在，尝试读取
                target = rel.get('Target')
                if target and target in zip_file.namelist():
                    print(f"  目标文件存在: {target}")
                    
                    # 如果是Visio文件
                    if target.endswith('.vsdx') or 'visio' in target.lower():
                        print("  这是一个Visio嵌入对象!")
                    elif target.endswith('.xlsx'):
                        print("  这是一个Excel嵌入对象!")
                    elif target.endswith('.pptx'):
                        print("  这是一个PowerPoint嵌入对象!")
                    else:
                        print(f"  嵌入对象类型: {target}")
                        
                break
                
    except Exception as e:
        print(f"  分析关系失败: {e}")

def extract_embedded_objects(docx_path):
    """提取嵌入对象"""
    print(f"\n=== 提取嵌入对象 ===")
    
    with zipfile.ZipFile(docx_path, 'r') as zip_file:
        # 查找embeddings目录
        embedding_files = [f for f in zip_file.namelist() if f.startswith('word/embeddings/')]
        
        print(f"找到 {len(embedding_files)} 个嵌入文件:")
        for emb_file in embedding_files:
            print(f"  {emb_file}")
            
            # 分析文件类型
            if emb_file.endswith('.vsdx'):
                print(f"    -> Visio图表文件")
                # 可以提取这个文件进行进一步分析
                analyze_visio_object(zip_file, emb_file)
            elif emb_file.endswith('.xlsx'):
                print(f"    -> Excel电子表格")
            elif emb_file.endswith('.pptx'):
                print(f"    -> PowerPoint演示文稿")
            else:
                print(f"    -> 其他类型: {emb_file}")

def analyze_visio_object(zip_file, visio_path):
    """分析Visio嵌入对象"""
    print(f"  分析Visio文件: {visio_path}")
    
    try:
        # 提取Visio文件
        visio_data = zip_file.read(visio_path)
        print(f"  文件大小: {len(visio_data)} 字节")
        
        # Visio文件也是ZIP格式，可以进一步解析
        # 但这里先只记录其存在
        print(f"  这是一个Visio绘图文件，包含流程图内容")
        
    except Exception as e:
        print(f"  读取Visio文件失败: {e}")

def main():
    docx_path = "Files/PLM2.0/BOM 审核申请.docx"
    
    if not os.path.exists(docx_path):
        print(f"文件不存在: {docx_path}")
        return
    
    print("开始分析DOCX文件中的嵌入对象...")
    
    # 1. 分析object元素
    analyze_object_element(docx_path)
    
    # 2. 提取嵌入对象
    extract_embedded_objects(docx_path)
    
    print("\n分析完成!")

if __name__ == "__main__":
    main()
