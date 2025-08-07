#!/usr/bin/env python3
"""
深度分析SmartArt XML结构，理解层次关系和逻辑顺序
"""

import zipfile
import os
from lxml import etree

def analyze_smartart_structure(docx_path):
    """深度分析SmartArt的XML结构"""
    
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # 首先读取文档关系
        rels_content = docx_zip.read('word/_rels/document.xml.rels')
        rels_tree = etree.fromstring(rels_content)
        
        # 命名空间
        namespaces = {
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        
        # 查找所有关系
        all_rels = rels_tree.xpath('.//r:Relationship', namespaces=namespaces)
        print(f"总共找到 {len(all_rels)} 个关系:")
        
        diagram_rels = []
        for rel in all_rels:
            rel_type = rel.get('Type', '')
            rel_id = rel.get('Id', '')
            target = rel.get('Target', '')
            print(f"  {rel_id}: {rel_type} -> {target}")
            
            if 'diagram' in rel_type.lower():
                diagram_rels.append(rel)
        
        print(f"\n找到 {len(diagram_rels)} 个diagram关系:")
        for rel in diagram_rels:
            print(f"  ID: {rel.get('Id')}, Type: {rel.get('Type')}, Target: {rel.get('Target')}")
        
        # 分析每个diagram文件
        for rel in diagram_rels:
            target = rel.get('Target')
            if target.startswith('diagrams/'):
                try:
                    file_content = docx_zip.read(f'word/{target}')
                    print(f"\n=== 分析文件: {target} ===")
                    analyze_diagram_xml(file_content, target)
                except Exception as e:
                    print(f"无法读取 {target}: {e}")

def analyze_diagram_xml(xml_content, filename):
    """分析diagram XML的详细结构"""
    
    try:
        tree = etree.fromstring(xml_content)
        print(f"根元素: {tree.tag}")
        
        # 定义更全面的命名空间
        namespaces = {
            'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
        
        if 'data' in filename:
            print("\n--- 数据模型分析 ---")
            analyze_data_model(tree, namespaces)
            print("\n--- 打印完整XML结构 ---")
            print(etree.tostring(tree, encoding='unicode', pretty_print=True))
        elif 'layout' in filename:
            print("\n--- 布局分析 ---")
            analyze_layout(tree, namespaces)
        elif 'colors' in filename:
            print("\n--- 颜色方案分析 ---")
        elif 'quickStyle' in filename:
            print("\n--- 样式分析 ---")
            
    except Exception as e:
        print(f"解析XML失败: {e}")

def analyze_data_model(tree, namespaces):
    """分析数据模型的层次结构"""
    
    # 查找所有点(节点)
    points = tree.xpath('.//dgm:pt', namespaces=namespaces)
    print(f"找到 {len(points)} 个数据点:")
    
    # 查找所有连接关系
    connections = tree.xpath('.//dgm:cxn', namespaces=namespaces)
    print(f"找到 {len(connections)} 个连接关系:")
    
    # 构建关系图
    print("\n连接关系详情:")
    for i, cxn in enumerate(connections):
        src_id = cxn.get('srcId')
        dest_id = cxn.get('destId')
        cxn_type = cxn.get('type', 'unknown')
        src_ord = cxn.get('srcOrd', '0')
        dest_ord = cxn.get('destOrd', '0')
        print(f"  连接 {i+1}: {src_id}[{src_ord}] -> {dest_id}[{dest_ord}] (类型: {cxn_type})")
    
    print("\n详细节点信息:")
    for i, pt in enumerate(points):
        pt_type = pt.get('type', 'unknown')
        model_id = pt.get('modelId', 'no-id')
        
        print(f"\n  点 {i+1}: type={pt_type}, modelId={model_id}")
        
        # 查找文本内容
        text_elements = pt.xpath('.//a:t', namespaces=namespaces)
        for j, text_elem in enumerate(text_elements):
            text_content = text_elem.text or ""
            print(f"    文本 {j+1}: '{text_content}'")
        
        # 查找这个节点的父子关系
        parent_connections = [c for c in connections if c.get('destId') == model_id]
        child_connections = [c for c in connections if c.get('srcId') == model_id]
        
        if parent_connections:
            parent_ids = [c.get('srcId') for c in parent_connections]
            print(f"    父节点: {parent_ids}")
        if child_connections:
            child_ids = [c.get('destId') for c in child_connections]
            print(f"    子节点: {child_ids}")

def analyze_layout(tree, namespaces):
    """分析布局结构"""
    
    layout_nodes = tree.xpath('.//dgm:layoutNode', namespaces=namespaces)
    print(f"找到 {len(layout_nodes)} 个布局节点")
    
    for i, node in enumerate(layout_nodes):
        name = node.get('name', 'unnamed')
        print(f"  布局节点 {i+1}: {name}")

if __name__ == "__main__":
    docx_path = "Files/PLM2.0/BOM 审核申请.docx"
    
    if os.path.exists(docx_path):
        print(f"分析文档: {docx_path}")
        analyze_smartart_structure(docx_path)
    else:
        print(f"文档不存在: {docx_path}")
