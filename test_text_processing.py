#!/usr/bin/env python3
"""
测试文本处理功能
根据需求文档规格说明书处理示例文档
"""

import sys
import os
import json
import logging

# 添加src目录到Python路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.parsers.document_parser import parse_docx
from src.processors.text_processor import process_document_to_text
from src.utils.text_utils import safe_filename

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_text_processing():
    """测试文本处理功能"""
    
    # 测试文件路径
    test_files = [
        "Files/examples/demo.docx",  # 示例文件
        "Files/PLM2.0/BOM 审核申请.docx"  # 实际文件
    ]
    
    for test_file in test_files:
        if not os.path.exists(test_file):
            logger.warning(f"测试文件不存在: {test_file}")
            continue
            
        logger.info(f"\n{'='*60}")
        logger.info(f"测试文件: {test_file}")
        logger.info(f"{'='*60}")
        
        try:
            # 创建输出目录
            file_basename = os.path.splitext(os.path.basename(test_file))[0]
            safe_name = safe_filename(file_basename)
            output_dir = f"test_output/{safe_name}"
            os.makedirs(output_dir, exist_ok=True)
            
            # 解析文档
            logger.info("开始解析文档...")
            document_structure = parse_docx(test_file, output_dir, quick_mode=True)
            
            if not document_structure:
                logger.error("文档解析失败")
                continue
            
            # 保存JSON结果
            json_path = os.path.join(output_dir, "document.json")
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(document_structure, f, ensure_ascii=False, indent=2)
            logger.info(f"JSON结果已保存到: {json_path}")
            
            # 处理为标准化文本
            logger.info("开始文本处理...")
            doc_name = safe_name
            processed_text = process_document_to_text(document_structure, doc_name)
            
            # 保存处理后的文本
            text_path = os.path.join(output_dir, "processed_text.txt")
            with open(text_path, "w", encoding="utf-8") as f:
                f.write(processed_text)
            
            logger.info(f"标准化文本已保存到: {text_path}")
            
            # 显示文本预览
            if processed_text:
                logger.info("文本预览（前500字符）:")
                logger.info("-" * 40)
                print(processed_text[:500] + "..." if len(processed_text) > 500 else processed_text)
                logger.info("-" * 40)
                logger.info(f"总字符数: {len(processed_text)}")
            else:
                logger.warning("处理后的文本为空")
            
            # 统计信息
            sections = document_structure.get("sections", [])
            images = document_structure.get("images", {})
            processing_info = document_structure.get("processing_info", {})
            
            logger.info(f"\n统计信息:")
            logger.info(f"  章节数量: {len(sections)}")
            logger.info(f"  图片数量: {len(images)}")
            logger.info(f"  警告数量: {len(processing_info.get('warnings', []))}")
            logger.info(f"  错误数量: {len(processing_info.get('errors', []))}")
            
        except Exception as e:
            logger.error(f"处理文件 {test_file} 时发生错误: {e}")
            import traceback
            logger.error(traceback.format_exc())

def print_separator(title):
    """打印分隔线"""
    print("\n" + "="*60)
    print(f"  {title}")
    print("="*60)

def main():
    """主函数"""
    print_separator("文本处理功能测试")
    
    # 创建测试输出目录
    os.makedirs("test_output", exist_ok=True)
    
    # 运行测试
    test_text_processing()
    
    print_separator("测试完成")
    print("📁 测试结果保存在 test_output/ 目录中")
    print("📄 查看 processed_text.txt 文件了解处理结果")

if __name__ == "__main__":
    main()
