#!/usr/bin/env python3
"""
DOCX解析器 - 模块化版本
高性能的Word文档解析工具，支持提取文本、图片、SmartArt和嵌入对象

功能特点:
- 模块化架构，易于维护和扩展
- 支持单文件和批量处理
- 智能内容去重和路径处理
- 标准化文本输出格式
- 自动默认路径处理

使用方法:
    # 使用默认示例文件
    python docx_parser_modular.py
    
    # 处理单个文件
    python docx_parser_modular.py Files/example.docx
    
    # 批量处理文件夹
    python docx_parser_modular.py Files/PLM2.0
    
    # 指定输出目录
    python docx_parser_modular.py Files/example.docx output_folder

版本: 2.0
作者: DOCX Parser Team
"""

import sys
import os
import json
import logging
from typing import Optional

# 添加src目录到Python路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.parsers.document_parser import parse_docx
from src.parsers.batch_processor import process_docx_folder
from src.processors.text_processor import process_document_to_text
from src.utils.text_utils import safe_filename

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("docx_parser.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def main() -> None:
    """
    主函数，处理命令行参数并执行相应的解析操作
    
    支持的参数模式:
    - 无参数: 使用默认示例文件
    - 1个参数: 输入文件/文件夹路径
    - 2个参数: 输入路径和输出目录
    """
    # 处理命令行参数
    if len(sys.argv) == 1:
        # 没有任何参数，使用默认路径
        input_path = "Files/examples"
        custom_output_dir = None
        print(f"未指定输入路径，使用默认路径: {input_path}")
    elif len(sys.argv) == 2:
        # 只给出输入路径，输出到默认的parsed_docs目录
        input_path = sys.argv[1]
        custom_output_dir = "parsed_docs"
        print(f"未指定输出路径，将保存到: {custom_output_dir}")
    elif len(sys.argv) == 3:
        # 给出输入和输出路径
        input_path = sys.argv[1]
        custom_output_dir = sys.argv[2]
    else:
        print("用法: python docx_parser_modular.py [输入路径] [输出路径]")
        print("示例: python docx_parser_modular.py                          # 默认处理 Files/examples")
        print("示例: python docx_parser_modular.py demo.docx                # 输出到 parsed_docs/")
        print("示例: python docx_parser_modular.py demo.docx my_output/     # 自定义输出路径")
        print("示例: python docx_parser_modular.py Files/PLM2.0/            # 批量处理到 parsed_docs/")
        sys.exit(1)
    
    # 检查输入路径是否存在
    if not os.path.exists(input_path):
        logger.error(f"输入路径不存在: {input_path}")
        sys.exit(1)
    
    # 判断是单个文件还是文件夹
    if os.path.isfile(input_path):
        # 处理单个文件
        if not input_path.lower().endswith('.docx'):
            logger.error(f"不是有效的DOCX文件: {input_path}")
            sys.exit(1)
        
        # 为单个文件创建输出目录
        file_basename = os.path.splitext(os.path.basename(input_path))[0]
        safe_name = safe_filename(file_basename)
        
        if custom_output_dir == "parsed_docs":
            # 只给出输入路径的情况，使用parsed_docs作为根目录
            output_dir = os.path.join("parsed_docs", safe_name)
        elif custom_output_dir:
            # 给出自定义输出路径
            output_dir = os.path.join(custom_output_dir, safe_name)
        else:
            # 完全默认的情况（这应该不会发生，因为我们上面设置了默认值）
            output_dir = f"parsed_docs/single_files/{safe_name}"
        
        os.makedirs(output_dir, exist_ok=True)
        
        logger.info(f"开始处理单个文件: {input_path}")
        logger.info(f"输出目录: {output_dir}")
        
        # 解析文档（使用快速模式）
        document_structure = parse_docx(input_path, output_dir, quick_mode=True)
        
        if document_structure:
            # 保存为JSON文件
            json_path = os.path.join(output_dir, "document.json")
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(document_structure, f, ensure_ascii=False, indent=2)
            
            # 处理为标准化文本格式
            try:
                # 从文件名提取文档名称
                doc_name = safe_name
                processed_text = process_document_to_text(document_structure, doc_name, output_dir)
                
                # 保存处理后的文本
                text_path = os.path.join(output_dir, "processed_text.txt")
                with open(text_path, "w", encoding="utf-8") as f:
                    f.write(processed_text)
                
                logger.info(f"标准化文本已保存到: {text_path}")
                
            except Exception as e:
                logger.error(f"文本处理失败: {e}")
            
            # 统计信息
            total_images = len(document_structure.get("images", {}))
            processing_info = document_structure.get("processing_info", {})
            warnings = len(processing_info.get("warnings", []))
            errors = len(processing_info.get("errors", []))
            
            logger.info(f"文件处理完成!")
            logger.info(f"输出位置: {output_dir}")
            logger.info(f"图片数量: {total_images}")
            logger.info(f"警告数量: {warnings}")
            logger.info(f"错误数量: {errors}")
            
            if warnings > 0 or errors > 0:
                logger.info("详细的警告和错误信息请查看 document.json 文件")
        else:
            logger.error("文件处理失败")
            sys.exit(1)
            
    elif os.path.isdir(input_path):
        # 批量处理文件夹
        # 创建与输入文件夹对应的输出目录
        folder_name = os.path.basename(input_path.rstrip('/'))
        
        if custom_output_dir == "parsed_docs":
            # 只给出输入路径的情况，使用parsed_docs作为根目录
            output_base_dir = os.path.join("parsed_docs", f"{folder_name}_parsed")
        elif custom_output_dir:
            # 给出自定义输出路径
            output_base_dir = os.path.join(custom_output_dir, f"{folder_name}_parsed")
        else:
            # 完全默认的情况（这应该不会发生，因为我们上面设置了默认值）
            output_base_dir = f"parsed_docs/{folder_name}_parsed"
        
        logger.info(f"开始批量处理文件夹: {input_path}")
        logger.info(f"输出目录: {output_base_dir}")
        
        # 处理所有DOCX文件
        processed_count = process_docx_folder(input_path, output_base_dir, quick_mode=True)
        
        # 打印总结报告
        if processed_count and processed_count > 0:
            summary_path = os.path.join(output_base_dir, "summary.json")
            if os.path.exists(summary_path):
                with open(summary_path, "r", encoding="utf-8") as f:
                    summary = json.load(f)
                    print("\n" + "="*60)
                    print("📊 处理总结报告")
                    print("="*60)
                    print(f"📁 处理文件夹: {summary['input_folder']}")
                    print(f"📄 文件总数: {summary['total_files']}")
                    print(f"✅ 成功处理: {summary['processed']}")
                    print(f"❌ 失败文件: {summary['failed']}")
                    print(f"📈 成功率: {summary.get('success_rate', 'N/A')}")
                    
                    if summary.get('failed_files'):
                        print("\n❌ 失败文件列表:")
                        for failed in summary['failed_files'][:20]:
                            print(f"   - {failed.get('file', 'Unknown')}: {failed.get('error', 'Unknown error')}")
                        if len(summary['failed_files']) > 20:
                            print(f"   ... 还有 {len(summary['failed_files'])-20} 个失败文件")

                    print(f"\n📁 结果保存在: {output_base_dir}")
                    print("="*60)
        else:
            logger.error("没有成功处理任何文件")
            sys.exit(1)
    else:
        logger.error(f"无法识别的输入路径类型: {input_path}")
        sys.exit(1)

if __name__ == "__main__":
    main()
