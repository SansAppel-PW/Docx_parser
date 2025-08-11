#!/usr/bin/env python3
"""
DOCX解析器 - 模块化版本
高性能的Word文档解析工具，支持提取文本、图片、SmartArt和嵌入对象

使用方法:
    # 处理单个文件
    python docx_parser_modular.py Files/example.docx
    
    # 批量处理文件夹
    python docx_parser_modular.py Files/PLM2.0
    
    # 快速测试
    python quick_parse_example.py
"""

import sys
import os
import json
import logging

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

def main():
    """主函数"""
    # 示例使用 - 可以处理单个文件或批量处理文件夹
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    else:
        input_path = "Files/PLM2.0"  # 默认DOCX文件夹路径
    
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
                processed_text = process_document_to_text(document_structure, doc_name)
                
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
        output_base_dir = f"parsed_docs/{folder_name}_parsed"  # 输出目录
        
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
