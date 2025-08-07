#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析DOCX文件中遗漏内容的脚本
用于找出解析器未识别的内容类型
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from lxml import etree
import json
import re
from docx import Document
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

def iter_block_items(parent):
    """按文档顺序生成段落和表格"""
    if isinstance(parent, _Document):
        parent_elem = parent.element.body
    elif isinstance(parent, Table):
        parent_elem = parent._tbl
    else:
        raise ValueError("Unsupported parent type")
    
    for child in parent_elem.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def analyze_docx_structure(docx_path):
    """分析DOCX文件的完整结构"""
    print(f"分析文件: {docx_path}")
    
    # 1. 使用zipfile直接分析文档结构
    print("\n=== 文档内部文件结构 ===")
    with zipfile.ZipFile(docx_path, 'r') as zip_file:
        file_list = zip_file.namelist()
        for file_name in sorted(file_list):
            print(f"  {file_name}")
    
    # 2. 分析document.xml的原始结构
    print("\n=== document.xml 结构分析 ===")
    with zipfile.ZipFile(docx_path, 'r') as zip_file:
        document_xml = zip_file.read('word/document.xml')
        
        # 解析XML
        root = etree.fromstring(document_xml)
        
        # 定义命名空间
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
        }
        
        # 查找body中的所有直接子元素
        body = root.find('.//w:body', namespaces)
        if body is not None:
            print("Body中的所有直接子元素:")
            for i, child in enumerate(body):
                tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                print(f"  {i:3d}: {tag_name}")
                
                # 如果是段落，检查其内容
                if tag_name == 'p':
                    # 获取段落文本
                    text_content = ""
                    text_elements = child.findall('.//w:t', namespaces)
                    for t_elem in text_elements:
                        if t_elem.text:
                            text_content += t_elem.text
                    
                    # 检查是否包含特殊元素
                    drawings = child.findall('.//w:drawing', namespaces)
                    objects = child.findall('.//w:object', namespaces)
                    picts = child.findall('.//w:pict', namespaces)
                    
                    print(f"       文本: {text_content[:50]}{'...' if len(text_content) > 50 else ''}")
                    if drawings:
                        print(f"       包含 {len(drawings)} 个 drawing 元素")
                    if objects:
                        print(f"       包含 {len(objects)} 个 object 元素")
                    if picts:
                        print(f"       包含 {len(picts)} 个 pict 元素")
    
    return namespaces

def find_target_content(docx_path):
    """查找目标内容周围的详细信息"""
    print("\n=== 查找目标内容 ===")
    
    # 目标文本
    target_text1 = "2.刷新随签子件节点根据主签审对象校验并查询子件"
    target_text2 = "Uat中工作流模板在 启动MBOM流程 后有表达式"
    
    with zipfile.ZipFile(docx_path, 'r') as zip_file:
        document_xml = zip_file.read('word/document.xml')
        root = etree.fromstring(document_xml)
        
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        }
        
        body = root.find('.//w:body', namespaces)
        found_first = False
        found_second = False
        target_start_idx = None
        target_end_idx = None
        
        for i, child in enumerate(body):
            tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            if tag_name == 'p':
                # 获取段落文本
                text_content = ""
                text_elements = child.findall('.//w:t', namespaces)
                for t_elem in text_elements:
                    if t_elem.text:
                        text_content += t_elem.text
                
                # 检查是否为目标文本
                if target_text1 in text_content:
                    print(f"找到第一个目标文本在位置 {i}")
                    found_first = True
                    target_start_idx = i
                
                elif target_text2 in text_content and found_first:
                    print(f"找到第二个目标文本在位置 {i}")
                    found_second = True
                    target_end_idx = i
                    break
        
        # 分析两个目标文本之间的内容
        if target_start_idx is not None and target_end_idx is not None:
            print(f"\n=== 分析位置 {target_start_idx} 到 {target_end_idx} 之间的内容 ===")
            
            for i in range(target_start_idx + 1, target_end_idx):
                child = body[i]
                tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                
                print(f"\n位置 {i}: {tag_name}")
                
                # 详细分析这个元素
                if tag_name == 'p':
                    analyze_paragraph_detailed(child, namespaces, i)
                elif tag_name == 'tbl':
                    print("  这是一个表格")
                else:
                    print(f"  未知元素类型: {tag_name}")
                    # 打印元素的XML内容
                    xml_str = etree.tostring(child, encoding='unicode', pretty_print=True)
                    print(f"  XML内容: {xml_str[:500]}...")

def analyze_paragraph_detailed(para_elem, namespaces, position):
    """详细分析段落元素"""
    print(f"  段落详细分析:")
    
    # 获取文本内容
    text_content = ""
    text_elements = para_elem.findall('.//w:t', namespaces)
    for t_elem in text_elements:
        if t_elem.text:
            text_content += t_elem.text
    
    print(f"    文本内容: '{text_content}'")
    
    # 检查各种子元素
    runs = para_elem.findall('.//w:r', namespaces)
    print(f"    包含 {len(runs)} 个 run 元素")
    
    for j, run in enumerate(runs):
        print(f"      Run {j}:")
        
        # 检查run中的各种元素
        run_text = ""
        run_text_elems = run.findall('.//w:t', namespaces)
        for t_elem in run_text_elems:
            if t_elem.text:
                run_text += t_elem.text
        
        if run_text:
            print(f"        文本: '{run_text}'")
        
        # 检查drawing元素
        drawings = run.findall('.//w:drawing', namespaces)
        if drawings:
            print(f"        包含 {len(drawings)} 个 drawing 元素")
            for k, drawing in enumerate(drawings):
                analyze_drawing_element(drawing, namespaces, k)
        
        # 检查object元素
        objects = run.findall('.//w:object', namespaces)
        if objects:
            print(f"        包含 {len(objects)} 个 object 元素")
            for k, obj in enumerate(objects):
                analyze_object_element(obj, namespaces, k)
        
        # 检查pict元素
        picts = run.findall('.//w:pict', namespaces)
        if picts:
            print(f"        包含 {len(picts)} 个 pict 元素")
        
        # 检查其他特殊元素
        special_elements = []
        for elem in run:
            elem_name = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if elem_name not in ['t', 'rPr', 'br', 'tab']:
                special_elements.append(elem_name)
        
        if special_elements:
            print(f"        特殊元素: {', '.join(special_elements)}")

def analyze_drawing_element(drawing, namespaces, index):
    """分析drawing元素"""
    print(f"          Drawing {index}:")
    
    # 检查inline或anchor
    inline = drawing.find('.//wp:inline', namespaces)
    anchor = drawing.find('.//wp:anchor', namespaces)
    
    if inline is not None:
        print(f"            类型: inline")
    elif anchor is not None:
        print(f"            类型: anchor")
    
    # 检查graphic元素
    graphics = drawing.findall('.//a:graphic', namespaces)
    for graphic in graphics:
        graphic_data = graphic.find('.//a:graphicData', namespaces)
        if graphic_data is not None:
            uri = graphic_data.get('uri', '')
            print(f"            GraphicData URI: {uri}")
            
            # 如果是图片
            if 'picture' in uri:
                print(f"              -> 这是一个图片")
            # 如果是SmartArt
            elif 'diagram' in uri:
                print(f"              -> 这是一个SmartArt图表")
            # 如果是图表
            elif 'chart' in uri:
                print(f"              -> 这是一个图表")
            else:
                print(f"              -> 未知类型的图形内容")

def analyze_object_element(obj, namespaces, index):
    """分析object元素"""
    print(f"          Object {index}:")
    
    # 获取object的XML内容
    xml_str = etree.tostring(obj, encoding='unicode', pretty_print=True)
    print(f"            XML: {xml_str[:200]}...")

def compare_with_parsed_result(docx_path):
    """与解析结果进行对比"""
    print("\n=== 与python-docx解析结果对比 ===")
    
    doc = Document(docx_path)
    
    target_text1 = "2.刷新随签子件节点根据主签审对象校验并查询子件"
    target_text2 = "Uat中工作流模板在 启动MBOM流程 后有表达式"
    
    found_first = False
    found_second = False
    between_content = []
    
    for i, block in enumerate(iter_block_items(doc)):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            
            if target_text1 in text:
                print(f"python-docx: 找到第一个目标文本在块 {i}")
                found_first = True
                continue
            
            if target_text2 in text and found_first:
                print(f"python-docx: 找到第二个目标文本在块 {i}")
                found_second = True
                break
            
            if found_first:
                between_content.append({
                    'type': 'paragraph',
                    'index': i,
                    'text': text,
                    'runs_count': len(block.runs) if hasattr(block, 'runs') else 0
                })
        
        elif isinstance(block, Table):
            if found_first and not found_second:
                between_content.append({
                    'type': 'table',
                    'index': i,
                    'rows': len(block.rows),
                    'cols': len(block.columns) if hasattr(block, 'columns') else 0
                })
    
    print(f"\npython-docx 在两个目标文本之间找到 {len(between_content)} 个内容块:")
    for content in between_content:
        if content['type'] == 'paragraph':
            print(f"  段落 {content['index']}: '{content['text'][:50]}...' (runs: {content['runs_count']})")
        elif content['type'] == 'table':
            print(f"  表格 {content['index']}: {content['rows']}行 x {content['cols']}列")

def main():
    docx_path = "Files/PLM2.0/BOM 审核申请.docx"
    
    if not os.path.exists(docx_path):
        print(f"文件不存在: {docx_path}")
        return
    
    print("开始分析DOCX文件中的遗漏内容...")
    
    # 1. 分析文档结构
    namespaces = analyze_docx_structure(docx_path)
    
    # 2. 查找目标内容
    find_target_content(docx_path)
    
    # 3. 与解析结果对比
    compare_with_parsed_result(docx_path)
    
    print("\n分析完成!")

if __name__ == "__main__":
    main()
