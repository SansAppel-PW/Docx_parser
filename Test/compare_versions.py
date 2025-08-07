#!/usr/bin/env python3
"""
对比原始版本和模块化版本的解析结果
"""

import os
import sys
import json
import time
import logging
import subprocess
from pathlib import Path

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def run_parser(script_name, input_file, output_suffix):
    """运行解析器并返回结果路径"""
    logger.info(f"🚀 运行 {script_name}...")
    
    # 修改脚本中的输出目录
    if script_name == "docx_parser.py":
        # 临时修改原始脚本的输出目录
        output_dir = f"comparison_test/original_{output_suffix}"
    else:
        # 对于模块化版本，直接修改输出
        output_dir = f"comparison_test/modular_{output_suffix}"
    
    # 设置环境变量来控制输出目录
    env = os.environ.copy()
    env['CUSTOM_OUTPUT_DIR'] = output_dir
    
    try:
        start_time = time.time()
        
        # 运行解析器
        result = subprocess.run([
            sys.executable, script_name, input_file
        ], capture_output=True, text=True, env=env, timeout=60)
        
        end_time = time.time()
        
        if result.returncode == 0:
            logger.info(f"✅ {script_name} 运行成功，耗时 {end_time - start_time:.2f}s")
            return output_dir
        else:
            logger.error(f"❌ {script_name} 运行失败:")
            logger.error(f"STDOUT: {result.stdout}")
            logger.error(f"STDERR: {result.stderr}")
            return None
            
    except subprocess.TimeoutExpired:
        logger.error(f"❌ {script_name} 运行超时")
        return None
    except Exception as e:
        logger.error(f"❌ {script_name} 运行异常: {e}")
        return None

def analyze_results(original_dir, modular_dir):
    """分析两个版本的结果差异"""
    logger.info("📊 分析解析结果差异...")
    
    comparison_report = {
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "original_dir": original_dir,
        "modular_dir": modular_dir,
        "differences": [],
        "summary": {}
    }
    
    # 检查目录是否存在
    if not os.path.exists(original_dir):
        logger.error(f"原始版本输出目录不存在: {original_dir}")
        return None
        
    if not os.path.exists(modular_dir):
        logger.error(f"模块化版本输出目录不存在: {modular_dir}")
        return None
    
    # 分析目录结构
    original_structure = get_directory_structure(original_dir)
    modular_structure = get_directory_structure(modular_dir)
    
    logger.info("📁 目录结构对比:")
    logger.info(f"原始版本: {list(original_structure.keys())}")
    logger.info(f"模块化版本: {list(modular_structure.keys())}")
    
    # 对比图片数量
    original_images = len([f for f in original_structure.get('images', []) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.tiff', '.emf'))])
    modular_images = len([f for f in modular_structure.get('images', []) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.tiff', '.emf'))])
    
    logger.info(f"🖼️  图片数量对比: 原始版本 {original_images} vs 模块化版本 {modular_images}")
    
    # 对比JSON文件
    original_json = os.path.join(original_dir, "document.json")
    modular_json = os.path.join(modular_dir, "document.json")
    
    original_data = None
    modular_data = None
    
    if os.path.exists(original_json):
        try:
            with open(original_json, 'r', encoding='utf-8') as f:
                original_data = json.load(f)
        except Exception as e:
            logger.error(f"读取原始版本JSON失败: {e}")
    
    if os.path.exists(modular_json):
        try:
            with open(modular_json, 'r', encoding='utf-8') as f:
                modular_data = json.load(f)
        except Exception as e:
            logger.error(f"读取模块化版本JSON失败: {e}")
    
    # 对比JSON内容
    if original_data and modular_data:
        compare_json_content(original_data, modular_data, comparison_report)
    
    # 填充汇总信息
    comparison_report["summary"] = {
        "original_images": original_images,
        "modular_images": modular_images,
        "image_difference": modular_images - original_images,
        "original_has_smartart": "smartart" in original_structure,
        "modular_has_smartart": "smartart" in modular_structure,
        "original_has_embedded": "embedded_objects" in original_structure,
        "modular_has_embedded": "embedded_objects" in modular_structure
    }
    
    return comparison_report

def get_directory_structure(directory):
    """获取目录结构"""
    structure = {}
    try:
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            if os.path.isdir(item_path):
                structure[item] = os.listdir(item_path)
            else:
                if "files" not in structure:
                    structure["files"] = []
                structure["files"].append(item)
    except Exception as e:
        logger.error(f"获取目录结构失败: {e}")
    
    return structure

def compare_json_content(original_data, modular_data, report):
    """对比JSON内容"""
    logger.info("📄 对比JSON内容...")
    
    # 对比图片数量
    original_images = len(original_data.get("images", {}))
    modular_images = len(modular_data.get("images", {}))
    
    if original_images != modular_images:
        report["differences"].append({
            "type": "image_count",
            "original": original_images,
            "modular": modular_images,
            "difference": modular_images - original_images
        })
    
    # 对比章节数量
    original_sections = len(original_data.get("sections", []))
    modular_sections = len(modular_data.get("sections", []))
    
    if original_sections != modular_sections:
        report["differences"].append({
            "type": "section_count",
            "original": original_sections,
            "modular": modular_sections,
            "difference": modular_sections - original_sections
        })
    
    # 对比处理信息
    original_info = original_data.get("processing_info", {})
    modular_info = modular_data.get("processing_info", {})
    
    for key in ["blocks_processed", "tables_found", "images_found"]:
        original_val = original_info.get(key, 0)
        modular_val = modular_info.get(key, 0)
        
        if original_val != modular_val:
            report["differences"].append({
                "type": key,
                "original": original_val,
                "modular": modular_val,
                "difference": modular_val - original_val
            })

def print_comparison_report(report):
    """打印对比报告"""
    print("\n" + "="*80)
    print("📊 DOCX解析器版本对比报告")
    print("="*80)
    print(f"⏰ 生成时间: {report['timestamp']}")
    print()
    
    summary = report["summary"]
    print("📈 关键指标对比:")
    print(f"   🖼️  图片数量: 原始版本 {summary['original_images']} → 模块化版本 {summary['modular_images']} (差异: {summary['image_difference']:+d})")
    print(f"   📊 SmartArt: 原始版本 {'✅' if summary['original_has_smartart'] else '❌'} → 模块化版本 {'✅' if summary['modular_has_smartart'] else '❌'}")
    print(f"   📎 嵌入对象: 原始版本 {'✅' if summary['original_has_embedded'] else '❌'} → 模块化版本 {'✅' if summary['modular_has_embedded'] else '❌'}")
    print()
    
    if report["differences"]:
        print("⚠️  发现的差异:")
        for diff in report["differences"]:
            print(f"   • {diff['type']}: {diff['original']} → {diff['modular']} (差异: {diff['difference']:+d})")
    else:
        print("✅ 未发现显著差异")
    
    print()
    print(f"📁 详细结果:")
    print(f"   原始版本: {report['original_dir']}")
    print(f"   模块化版本: {report['modular_dir']}")
    print("="*80)

def main():
    """主函数"""
    test_file = "Files/PLM2.0/BOM 审核申请.docx"
    
    if not os.path.exists(test_file):
        logger.error(f"测试文件不存在: {test_file}")
        return
    
    logger.info("🔍 开始DOCX解析器版本对比测试")
    logger.info(f"📄 测试文件: {test_file}")
    
    # 运行原始版本（临时修改输出目录）
    logger.info("\n" + "="*50)
    
    # 直接运行并手动设置输出目录
    original_dir = "comparison_test/original_result"
    modular_dir = "comparison_test/modular_result"
    
    # 先运行原始版本
    logger.info("🚀 运行原始版本...")
    try:
        # 临时备份并修改原始脚本
        with open("docx_parser.py", "r", encoding="utf-8") as f:
            original_content = f.read()
        
        # 修改输出目录
        modified_content = original_content.replace(
            'output_dir = f"parsed_docs/single_files/{safe_name}"',
            f'output_dir = "{original_dir}"'
        )
        
        with open("docx_parser_temp.py", "w", encoding="utf-8") as f:
            f.write(modified_content)
        
        # 运行修改后的版本
        result = subprocess.run([sys.executable, "docx_parser_temp.py", test_file], 
                              capture_output=True, text=True, timeout=60)
        
        # 清理临时文件
        os.remove("docx_parser_temp.py")
        
        if result.returncode == 0:
            logger.info("✅ 原始版本运行成功")
        else:
            logger.error(f"❌ 原始版本运行失败: {result.stderr}")
            return
            
    except Exception as e:
        logger.error(f"❌ 运行原始版本失败: {e}")
        return
    
    # 运行模块化版本
    logger.info("🚀 运行模块化版本...")
    try:
        # 临时修改模块化版本的输出目录
        with open("docx_parser_modular.py", "r", encoding="utf-8") as f:
            modular_content = f.read()
        
        modified_modular = modular_content.replace(
            'output_dir = f"parsed_docs/single_files/{safe_name}"',
            f'output_dir = "{modular_dir}"'
        )
        
        with open("docx_parser_modular_temp.py", "w", encoding="utf-8") as f:
            f.write(modified_modular)
        
        result = subprocess.run([sys.executable, "docx_parser_modular_temp.py", test_file], 
                              capture_output=True, text=True, timeout=60)
        
        # 清理临时文件
        os.remove("docx_parser_modular_temp.py")
        
        if result.returncode == 0:
            logger.info("✅ 模块化版本运行成功")
        else:
            logger.error(f"❌ 模块化版本运行失败: {result.stderr}")
            return
            
    except Exception as e:
        logger.error(f"❌ 运行模块化版本失败: {e}")
        return
    
    # 分析结果
    logger.info("\n📊 分析结果差异...")
    report = analyze_results(original_dir, modular_dir)
    
    if report:
        print_comparison_report(report)
        
        # 保存报告
        report_file = "comparison_test/comparison_report.json"
        with open(report_file, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        
        logger.info(f"\n📄 详细报告已保存到: {report_file}")
    
    logger.info("\n🎉 对比测试完成!")

if __name__ == "__main__":
    main()
