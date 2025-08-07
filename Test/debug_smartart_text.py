#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
改进的SmartArt文本提取脚本
"""

import os
import zipfile
from lxml import etree
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def debug_smartart_content(docx_path):
    """
    调试SmartArt内容，查看原始XML结构
    """
    logger.info(f"调试SmartArt内容: {docx_path}")
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            # 列出所有diagram相关文件
            diagram_files = [f for f in docx_zip.namelist() if 'diagram' in f.lower()]
            logger.info(f"发现diagram文件: {diagram_files}")
            
            for diagram_file in diagram_files:
                logger.info(f"\n分析文件: {diagram_file}")
                logger.info("=" * 50)
                
                try:
                    with docx_zip.open(diagram_file) as f:
                        content = f.read().decode('utf-8')
                        
                        # 美化XML输出
                        try:
                            root = etree.fromstring(content.encode('utf-8'))
                            pretty_xml = etree.tostring(root, encoding='unicode', pretty_print=True)
                            
                            # 只显示前500个字符，避免输出过长
                            if len(pretty_xml) > 500:
                                logger.info(f"内容预览:\n{pretty_xml[:500]}...")
                            else:
                                logger.info(f"完整内容:\n{pretty_xml}")
                                
                        except Exception as e:
                            logger.error(f"XML解析失败: {e}")
                            # 显示原始内容的前500个字符
                            if len(content) > 500:
                                logger.info(f"原始内容预览:\n{content[:500]}...")
                            else:
                                logger.info(f"原始内容:\n{content}")
                
                except Exception as e:
                    logger.error(f"读取文件 {diagram_file} 失败: {e}")
    
    except Exception as e:
        logger.error(f"调试失败: {e}")

def extract_text_from_diagram_data(data_xml):
    """
    从diagram数据XML中提取所有可能的文本
    """
    texts = []
    try:
        # 尝试多种可能的命名空间组合
        possible_namespaces = [
            {
                'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            },
            {
                'dgm': 'http://schemas.microsoft.com/office/drawing/2008/diagram',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            }
        ]
        
        root = etree.fromstring(data_xml.encode('utf-8'))
        
        # 尝试不同的命名空间
        for namespaces in possible_namespaces:
            # 查找文本元素的多种可能路径
            text_paths = [
                './/a:t',           # 直接文本
                './/dgm:pt//a:t',   # 点中的文本
                './/text()',        # 所有文本节点
                './/*[local-name()="t"]',  # 忽略命名空间的文本
                './/*[contains(local-name(), "text")]'  # 包含text的元素
            ]
            
            for path in text_paths:
                try:
                    if path == './/text()':
                        # 获取所有文本节点
                        text_nodes = root.xpath(path)
                        for text_node in text_nodes:
                            text_content = text_node.strip()
                            if text_content and len(text_content) > 1:  # 排除单个字符
                                texts.append(text_content)
                    else:
                        # 获取元素
                        elements = root.findall(path, namespaces) if 'dgm:' in path or 'a:' in path else root.xpath(path)
                        for elem in elements:
                            text_content = elem.text if hasattr(elem, 'text') else str(elem)
                            if text_content:
                                text_content = text_content.strip()
                                if text_content and len(text_content) > 1:
                                    texts.append(text_content)
                except Exception as e:
                    # 忽略查找失败，继续尝试其他路径
                    pass
        
        # 去重并过滤
        unique_texts = []
        for text in texts:
            if text not in unique_texts and text.strip():
                unique_texts.append(text.strip())
        
        return unique_texts
        
    except Exception as e:
        logger.error(f"提取diagram文本失败: {e}")
        return []

def test_improved_text_extraction():
    """
    测试改进的文本提取
    """
    test_file = "Files/PLM2.0/BOM 审核申请.docx"
    
    if not os.path.exists(test_file):
        logger.error(f"测试文件不存在: {test_file}")
        return
    
    logger.info("开始调试SmartArt内容")
    debug_smartart_content(test_file)
    
    logger.info("\n开始提取文本")
    logger.info("=" * 50)
    
    try:
        with zipfile.ZipFile(test_file, 'r') as docx_zip:
            # 查找数据文件
            data_files = [f for f in docx_zip.namelist() if 'diagrams/data' in f]
            
            for data_file in data_files:
                logger.info(f"\n提取文件 {data_file} 中的文本:")
                
                with docx_zip.open(data_file) as f:
                    data_xml = f.read().decode('utf-8')
                    texts = extract_text_from_diagram_data(data_xml)
                    
                    if texts:
                        logger.info(f"发现 {len(texts)} 个文本项:")
                        for i, text in enumerate(texts, 1):
                            logger.info(f"  {i}. {text}")
                    else:
                        logger.info("未发现文本内容")
    
    except Exception as e:
        logger.error(f"测试失败: {e}")

if __name__ == "__main__":
    test_improved_text_extraction()
