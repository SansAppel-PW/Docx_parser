#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
专门检测和测试 SmartArt 内容的脚本
"""

import os
import zipfile
from docx import Document
from lxml import etree
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def find_smartart_detailed(docx_path):
    """
    详细查找SmartArt内容
    """
    logger.info(f"详细分析文档: {docx_path}")
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            file_list = docx_zip.namelist()
            
            # 查找diagram相关文件
            diagram_files = [f for f in file_list if 'diagram' in f.lower()]
            if diagram_files:
                logger.info(f"发现diagram文件: {diagram_files}")
                return True
            
            # 检查主文档内容
            with docx_zip.open('word/document.xml') as doc_xml:
                content = doc_xml.read().decode('utf-8')
                
                # 解析XML以查找SmartArt
                root = etree.fromstring(content.encode('utf-8'))
                
                namespaces = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
                }
                
                # 查找所有graphic元素
                graphics = root.findall('.//a:graphic', namespaces)
                logger.info(f"发现 {len(graphics)} 个graphic元素")
                
                smartart_found = False
                for i, graphic in enumerate(graphics):
                    graphic_data = graphic.find('.//a:graphicData', namespaces)
                    if graphic_data is not None:
                        uri = graphic_data.get('uri', '')
                        logger.info(f"Graphic {i+1} URI: {uri}")
                        
                        # 检查是否为SmartArt
                        if 'diagram' in uri:
                            logger.info(f">>> 发现SmartArt图表!")
                            smartart_found = True
                            
                            # 尝试获取更多信息
                            rel_ids = graphic_data.findall('.//*')
                            for rel_id in rel_ids:
                                logger.info(f"  元素: {rel_id.tag} - {rel_id.attrib}")
                
                # 查找AlternateContent中可能包含的SmartArt
                alt_contents = root.findall('.//mc:AlternateContent', namespaces)
                for alt_content in alt_contents:
                    # 在choice中查找diagram相关内容
                    choices = alt_content.findall('.//mc:Choice', namespaces)
                    for choice in choices:
                        choice_graphics = choice.findall('.//a:graphic', namespaces)
                        for graphic in choice_graphics:
                            graphic_data = graphic.find('.//a:graphicData', namespaces)
                            if graphic_data is not None:
                                uri = graphic_data.get('uri', '')
                                if 'diagram' in uri:
                                    logger.info(f">>> 在AlternateContent/Choice中发现SmartArt!")
                                    smartart_found = True
                
                return smartart_found
                
    except Exception as e:
        logger.error(f"分析失败: {e}")
        return False

def create_test_smartart_doc():
    """
    创建一个包含SmartArt的测试文档
    """
    # 这里我们只是检查现有文档，不创建新文档
    pass

def test_all_documents():
    """
    测试所有可用的文档
    """
    test_dir = "Files/PLM2.0"
    if not os.path.exists(test_dir):
        logger.error(f"测试目录不存在: {test_dir}")
        return
    
    docx_files = [f for f in os.listdir(test_dir) if f.lower().endswith('.docx') and not f.startswith('~$')]
    
    logger.info(f"发现 {len(docx_files)} 个文档文件")
    
    smartart_docs = []
    
    for docx_file in docx_files:
        docx_path = os.path.join(test_dir, docx_file)
        logger.info(f"\n分析: {docx_file}")
        
        has_smartart = find_smartart_detailed(docx_path)
        if has_smartart:
            smartart_docs.append(docx_file)
            logger.info(f"✅ {docx_file} 包含SmartArt")
        else:
            logger.info(f"❌ {docx_file} 不包含SmartArt")
    
    logger.info(f"\n总结：")
    logger.info(f"包含SmartArt的文档 ({len(smartart_docs)}):")
    for doc in smartart_docs:
        logger.info(f"  - {doc}")

if __name__ == "__main__":
    test_all_documents()
