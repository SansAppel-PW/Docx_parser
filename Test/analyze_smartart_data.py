#!/usr/bin/env python3
"""
分析SmartArt数据文件内容
"""

import zipfile
import os
from lxml import etree

def analyze_smartart_data(docx_path):
    """分析SmartArt数据文件的详细内容"""
    
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # 读取SmartArt数据文件
        data_content = docx_zip.read('word/diagrams/data1.xml')
        print(f"SmartArt数据文件内容 ({len(data_content)} 字节):")
        
        # 解析XML
        tree = etree.fromstring(data_content)
        
        # 命名空间
        namespaces = {
            'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        
        print(f"根元素: {tree.tag}")
        
        # 查找所有数据点
        points = tree.xpath('.//dgm:pt', namespaces=namespaces)
        print(f"\n找到 {len(points)} 个数据点:")
        
        # 查找所有连接关系
        connections = tree.xpath('.//dgm:cxn', namespaces=namespaces)
        print(f"找到 {len(connections)} 个连接关系:")
        
        # 构建连接关系映射
        print("\n=== 连接关系详情 ===")
        for i, cxn in enumerate(connections):
            src_id = cxn.get('srcId')
            dest_id = cxn.get('destId')
            cxn_type = cxn.get('type', 'unknown')
            src_ord = cxn.get('srcOrd', '0')
            dest_ord = cxn.get('destOrd', '0')
            print(f"连接 {i+1}: {src_id}[{src_ord}] --{cxn_type}--> {dest_id}[{dest_ord}]")
        
        # 分析每个数据点
        print("\n=== 数据点详情 ===")
        for i, pt in enumerate(points):
            model_id = pt.get('modelId')
            pt_type = pt.get('type', 'unknown')
            
            print(f"\n数据点 {i+1}: {model_id} (类型: {pt_type})")
            
            # 查找文本内容
            text_elements = pt.xpath('.//a:t', namespaces=namespaces)
            if text_elements:
                for j, text_elem in enumerate(text_elements):
                    text_content = text_elem.text or ""
                    print(f"  文本 {j+1}: '{text_content}'")
            else:
                print("  无文本内容")
            
            # 查找父子关系
            parent_connections = [c for c in connections if c.get('destId') == model_id]
            child_connections = [c for c in connections if c.get('srcId') == model_id]
            
            if parent_connections:
                parent_info = []
                for pc in parent_connections:
                    parent_info.append(f"{pc.get('srcId')}[{pc.get('type', 'unknown')}]")
                print(f"  父节点: {', '.join(parent_info)}")
                
            if child_connections:
                child_info = []
                for cc in child_connections:
                    child_info.append(f"{cc.get('destId')}[{cc.get('type', 'unknown')}]")
                print(f"  子节点: {', '.join(child_info)}")
        
        # 构建层次结构树
        print("\n=== 层次结构分析 ===")
        build_hierarchy_tree(points, connections)
        
        # 显示原始XML结构
        print("\n=== 原始XML结构 ===")
        print(etree.tostring(tree, encoding='unicode', pretty_print=True))

def build_hierarchy_tree(points, connections):
    """构建并显示层次结构树"""
    
    # 创建点ID到点对象的映射
    points_map = {pt.get('modelId'): pt for pt in points}
    
    # 找到根节点(没有父节点的节点)
    all_dest_ids = {c.get('destId') for c in connections}
    all_src_ids = {c.get('srcId') for c in connections}
    
    root_ids = []
    for pt in points:
        pt_id = pt.get('modelId')
        if pt_id not in all_dest_ids:
            root_ids.append(pt_id)
    
    print(f"找到 {len(root_ids)} 个根节点: {root_ids}")
    
    # 构建父子关系映射
    children_map = {}
    for conn in connections:
        src_id = conn.get('srcId')
        dest_id = conn.get('destId')
        conn_type = conn.get('type')
        
        if src_id not in children_map:
            children_map[src_id] = []
        children_map[src_id].append({
            'id': dest_id,
            'type': conn_type,
            'srcOrd': conn.get('srcOrd', '0'),
            'destOrd': conn.get('destOrd', '0')
        })
    
    # 排序子节点
    for src_id in children_map:
        children_map[src_id].sort(key=lambda x: int(x.get('destOrd', '0')))
    
    # 递归显示树结构
    def print_tree(node_id, level=0, prefix=""):
        if node_id not in points_map:
            return
            
        pt = points_map[node_id]
        indent = "  " * level
        
        # 获取文本内容
        namespaces = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
        text_elements = pt.xpath('.//a:t', namespaces=namespaces)
        text_content = []
        for text_elem in text_elements:
            if text_elem.text:
                text_content.append(text_elem.text)
        
        text_str = " | ".join(text_content) if text_content else "[无文本]"
        print(f"{indent}{prefix}{node_id}: {text_str}")
        
        # 递归显示子节点
        if node_id in children_map:
            children = children_map[node_id]
            for i, child in enumerate(children):
                child_prefix = f"└─ [{child['type']}] " if i == len(children) - 1 else f"├─ [{child['type']}] "
                print_tree(child['id'], level + 1, child_prefix)
    
    # 从每个根节点开始显示
    for root_id in root_ids:
        print_tree(root_id)

if __name__ == "__main__":
    docx_path = "Files/PLM2.0/BOM 审核申请.docx"
    
    if os.path.exists(docx_path):
        print(f"分析文档: {docx_path}")
        analyze_smartart_data(docx_path)
    else:
        print(f"文档不存在: {docx_path}")
