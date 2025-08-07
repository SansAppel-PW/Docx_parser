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
        
        # 查找diagram关系
        diagram_rels = rels_tree.xpath('.//r:Relationship[contains(@Type, "diagram")]', namespaces=namespaces)
        
        print(f"找到 {len(diagram_rels)} 个diagram关系:")
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
    for i, cxn in enumerate(connections):
        src_id = cxn.get('srcId')
        dest_id = cxn.get('destId')
        cxn_type = cxn.get('type', 'unknown')
        print(f"  连接 {i+1}: {src_id} -> {dest_id} (类型: {cxn_type})")
    
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
            print(f"    父节点: {[c.get('srcId') for c in parent_connections]}")
        if child_connections:
            print(f"    子节点: {[c.get('destId') for c in child_connections]}")

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
    test_file = "Files/PLM2.0/BOM 审核申请.docx"
    
    if not os.path.exists(test_file):
        logger.error(f"测试文件不存在: {test_file}")
        return
    
    logger.info("分析SmartArt XML结构")
    logger.info("=" * 50)
    
    # 1. 分析文档主要XML中的SmartArt引用
    try:
        doc = Document(test_file)
        
        # 找到包含SmartArt的段落
        for para_idx, para in enumerate(doc.paragraphs):
            for run_idx, run in enumerate(para.runs):
                if run._element and run._element.xml:
                    xml_content = run._element.xml
                    if 'diagram' in xml_content:
                        logger.info(f"\n在段落{para_idx}运行{run_idx}中发现SmartArt引用")
                        logger.info("XML内容:")
                        
                        # 解析XML并美化输出
                        try:
                            root = etree.fromstring(xml_content)
                            pretty_xml = etree.tostring(root, encoding='unicode', pretty_print=True)
                            logger.info(pretty_xml)
                            
                            # 查找所有可能的关系引用
                            namespaces = {
                                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                                'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
                            }
                            
                            # 查找所有关系引用
                            rel_refs = root.findall('.//*[@r:id]', namespaces)
                            logger.info(f"发现关系引用: {len(rel_refs)}")
                            for ref in rel_refs:
                                rel_id = ref.get(f'{{{namespaces["r"]}}}id')
                                logger.info(f"  关系ID: {rel_id}, 元素: {ref.tag}")
                                
                            # 查找graphic元素
                            graphics = root.findall('.//a:graphic', namespaces)
                            for graphic in graphics:
                                graphic_data = graphic.find('.//a:graphicData', namespaces)
                                if graphic_data is not None:
                                    uri = graphic_data.get('uri', '')
                                    logger.info(f"GraphicData URI: {uri}")
                                    
                                    # 打印所有子元素
                                    for child in graphic_data:
                                        logger.info(f"  子元素: {child.tag}, 属性: {child.attrib}")
                                        
                                        # 进一步分析dgm元素
                                        if 'dgm' in child.tag:
                                            logger.info(f"  详细内容: {etree.tostring(child, encoding='unicode')}")
                        
                        except Exception as e:
                            logger.error(f"XML解析失败: {e}")
                            logger.info(f"原始XML内容:\n{xml_content}")
    
    except Exception as e:
        logger.error(f"分析文档失败: {e}")
    
    # 2. 直接分析ZIP文件中的关系文件
    logger.info("\n分析关系文件")
    logger.info("=" * 30)
    
    try:
        with zipfile.ZipFile(test_file, 'r') as docx_zip:
            # 查找所有.rels文件
            rels_files = [f for f in docx_zip.namelist() if f.endswith('.rels')]
            logger.info(f"发现关系文件: {rels_files}")
            
            for rels_file in rels_files:
                if 'word' in rels_file:  # 只分析word相关的关系
                    logger.info(f"\n分析关系文件: {rels_file}")
                    
                    with docx_zip.open(rels_file) as f:
                        content = f.read().decode('utf-8')
                        
                        # 查找diagram相关的关系
                        if 'diagram' in content:
                            logger.info("发现diagram相关关系:")
                            
                            root = etree.fromstring(content.encode('utf-8'))
                            relationships = root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
                            
                            for rel in relationships:
                                target = rel.get('Target', '')
                                rel_id = rel.get('Id', '')
                                rel_type = rel.get('Type', '')
                                
                                if 'diagram' in target.lower() or 'diagram' in rel_type.lower():
                                    logger.info(f"  ID: {rel_id}, Target: {target}, Type: {rel_type}")
    
    except Exception as e:
        logger.error(f"分析ZIP文件失败: {e}")

if __name__ == "__main__":
    analyze_smartart_structure()
