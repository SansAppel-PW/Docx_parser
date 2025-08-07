#!/usr/bin/env python3
"""
快速模式测试脚本
测试优化后的DOCX解析器性能
"""

import os
import time
import logging
from docx_parser import parse_docx, process_docx_folder

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_single_file_performance(docx_path, output_dir):
    """测试单个文件的解析性能"""
    if not os.path.exists(docx_path):
        logger.error(f"测试文件不存在: {docx_path}")
        return
    
    logger.info(f"开始测试文件: {os.path.basename(docx_path)}")
    
    # 测试快速模式
    logger.info("=== 快速模式测试 ===")
    start_time = time.time()
    quick_output = os.path.join(output_dir, "quick_mode")
    result_quick = parse_docx(docx_path, quick_output, quick_mode=True)
    quick_time = time.time() - start_time
    
    logger.info(f"快速模式耗时: {quick_time:.2f} 秒")
    
    # 测试标准模式（如果时间允许）
    if quick_time < 30:  # 只有快速模式小于30秒才测试标准模式
        logger.info("=== 标准模式测试 ===")
        start_time = time.time()
        standard_output = os.path.join(output_dir, "standard_mode")
        result_standard = parse_docx(docx_path, standard_output, quick_mode=False)
        standard_time = time.time() - start_time
        
        logger.info(f"标准模式耗时: {standard_time:.2f} 秒")
        logger.info(f"性能提升: {((standard_time - quick_time) / standard_time * 100):.1f}%")
    else:
        logger.info("快速模式已经很慢，跳过标准模式测试")
    
    # 分析结果
    if result_quick:
        images_count = len(result_quick.get("images", {}))
        smartart_count = sum(1 for section in result_quick.get("sections", []) 
                           for item in section.get("content", []) 
                           if item.get("type") == "smartart")
        embedded_count = sum(1 for section in result_quick.get("sections", []) 
                           for item in section.get("content", []) 
                           if item.get("type") == "embedded_object")
        
        logger.info(f"解析结果: 图片={images_count}, SmartArt={smartart_count}, 嵌入对象={embedded_count}")
        
        # 检查是否有处理错误
        processing_info = result_quick.get("processing_info", {})
        errors = processing_info.get("errors", [])
        warnings = processing_info.get("warnings", [])
        
        if errors:
            logger.warning(f"处理错误 ({len(errors)}): {errors[:3]}...")
        if warnings:
            logger.info(f"处理警告 ({len(warnings)}): {warnings[:2]}...")
    else:
        logger.error("解析失败")

def test_folder_performance(input_folder, output_dir):
    """测试文件夹批量处理性能"""
    if not os.path.exists(input_folder):
        logger.error(f"输入文件夹不存在: {input_folder}")
        return
    
    # 获取所有DOCX文件
    docx_files = [f for f in os.listdir(input_folder) 
                  if f.lower().endswith('.docx') and not f.startswith('~$')]
    
    if not docx_files:
        logger.warning(f"在 {input_folder} 中没有找到DOCX文件")
        return
    
    logger.info(f"找到 {len(docx_files)} 个DOCX文件")
    
    # 测试前几个文件
    test_files = docx_files[:3]  # 只测试前3个文件
    
    total_time = 0
    successful_files = 0
    
    for i, filename in enumerate(test_files, 1):
        docx_path = os.path.join(input_folder, filename)
        file_output = os.path.join(output_dir, f"file_{i}")
        os.makedirs(file_output, exist_ok=True)
        
        logger.info(f"测试文件 {i}/{len(test_files)}: {filename}")
        
        start_time = time.time()
        try:
            result = parse_docx(docx_path, file_output, quick_mode=True)
            if result:
                successful_files += 1
                file_time = time.time() - start_time
                total_time += file_time
                logger.info(f"  完成，耗时: {file_time:.2f} 秒")
            else:
                logger.warning(f"  解析失败")
        except Exception as e:
            logger.error(f"  解析出错: {e}")
    
    if successful_files > 0:
        avg_time = total_time / successful_files
        logger.info(f"\n性能统计:")
        logger.info(f"成功处理: {successful_files}/{len(test_files)} 个文件")
        logger.info(f"总耗时: {total_time:.2f} 秒")
        logger.info(f"平均每文件: {avg_time:.2f} 秒")
        logger.info(f"预估处理 {len(docx_files)} 个文件需要: {avg_time * len(docx_files):.1f} 秒")

if __name__ == "__main__":
    # 创建测试输出目录
    test_output_dir = "test_performance_output"
    os.makedirs(test_output_dir, exist_ok=True)
    
    # 测试选项
    test_type = input("选择测试类型 (1-单文件, 2-文件夹, 3-两者): ").strip()
    
    if test_type in ['1', '3']:
        # 单文件测试
        docx_file = input("输入DOCX文件路径 (回车使用默认): ").strip()
        if not docx_file:
            # 尝试找到测试文件
            possible_files = [
                "Files/PLM2.0/CBB新建申请V2.0.docx",
                "Files/PLM2.0/新ECR.docx",
                "Files/PLM2.0/材料维护申请流程V2.0.docx"
            ]
            for f in possible_files:
                if os.path.exists(f):
                    docx_file = f
                    break
        
        if docx_file and os.path.exists(docx_file):
            single_output = os.path.join(test_output_dir, "single_file_test")
            os.makedirs(single_output, exist_ok=True)
            test_single_file_performance(docx_file, single_output)
        else:
            logger.error("未找到有效的测试文件")
    
    if test_type in ['2', '3']:
        # 文件夹测试
        folder_path = input("输入文件夹路径 (回车使用默认 Files/PLM2.0): ").strip()
        if not folder_path:
            folder_path = "Files/PLM2.0"
        
        if os.path.exists(folder_path):
            folder_output = os.path.join(test_output_dir, "folder_test")
            os.makedirs(folder_output, exist_ok=True)
            test_folder_performance(folder_path, folder_output)
        else:
            logger.error(f"文件夹不存在: {folder_path}")
    
    logger.info("测试完成！")
