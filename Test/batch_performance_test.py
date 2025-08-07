#!/usr/bin/env python3
"""
批量测试对比两个版本的性能
"""

import os
import sys
import time
import json
import subprocess
import logging

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def run_batch_test(script_name, input_folder, output_suffix):
    """运行批量测试"""
    logger.info(f"🚀 运行批量测试: {script_name}")
    
    start_time = time.time()
    
    try:
        # 运行脚本
        result = subprocess.run([
            sys.executable, script_name, input_folder
        ], capture_output=True, text=True, timeout=300)  # 5分钟超时
        
        end_time = time.time()
        elapsed_time = end_time - start_time
        
        if result.returncode == 0:
            logger.info(f"✅ {script_name} 批量测试成功，耗时 {elapsed_time:.2f}s")
            
            # 解析输出中的统计信息
            output_lines = result.stdout.split('\n')
            stats = extract_stats_from_output(output_lines)
            stats['elapsed_time'] = elapsed_time
            
            return stats
        else:
            logger.error(f"❌ {script_name} 批量测试失败:")
            logger.error(f"STDERR: {result.stderr}")
            return None
            
    except subprocess.TimeoutExpired:
        logger.error(f"❌ {script_name} 批量测试超时")
        return None
    except Exception as e:
        logger.error(f"❌ {script_name} 批量测试异常: {e}")
        return None

def extract_stats_from_output(output_lines):
    """从输出中提取统计信息"""
    stats = {
        'total_files': 0,
        'processed': 0,
        'failed': 0,
        'success_rate': 0.0
    }
    
    for line in output_lines:
        if '文件总数:' in line:
            try:
                stats['total_files'] = int(line.split(':')[1].strip())
            except:
                pass
        elif '成功处理:' in line:
            try:
                stats['processed'] = int(line.split(':')[1].split()[0])
            except:
                pass
        elif '失败文件:' in line:
            try:
                stats['failed'] = int(line.split(':')[1].strip())
            except:
                pass
        elif '成功率:' in line:
            try:
                rate_str = line.split(':')[1].strip().replace('%', '')
                stats['success_rate'] = float(rate_str)
            except:
                pass
    
    return stats

def compare_batch_results():
    """对比批量处理结果"""
    logger.info("📊 开始批量处理对比测试")
    
    input_folder = "Files/PLM2.0"
    
    if not os.path.exists(input_folder):
        logger.error(f"测试文件夹不存在: {input_folder}")
        return
    
    # 清理现有结果
    if os.path.exists("parsed_docs"):
        import shutil
        shutil.rmtree("parsed_docs")
    
    # 运行原始版本
    logger.info("\n" + "="*60)
    logger.info("🔍 测试原始版本批量处理性能")
    original_stats = run_batch_test("docx_parser.py", input_folder, "original")
    
    # 稍等一下让文件系统稳定
    time.sleep(2)
    
    # 清理结果
    if os.path.exists("parsed_docs"):
        import shutil
        shutil.rmtree("parsed_docs")
    
    # 运行模块化版本
    logger.info("\n" + "="*60)
    logger.info("🔍 测试模块化版本批量处理性能")
    modular_stats = run_batch_test("docx_parser_modular.py", input_folder, "modular")
    
    # 生成对比报告
    if original_stats and modular_stats:
        print_batch_comparison(original_stats, modular_stats)
    else:
        logger.error("❌ 批量测试失败，无法生成对比报告")

def print_batch_comparison(original_stats, modular_stats):
    """打印批量处理对比报告"""
    print("\n" + "="*80)
    print("📊 批量处理性能对比报告")
    print("="*80)
    
    print("⏱️  处理时间对比:")
    print(f"   原始版本: {original_stats['elapsed_time']:.2f}s")
    print(f"   模块化版本: {modular_stats['elapsed_time']:.2f}s")
    
    time_diff = modular_stats['elapsed_time'] - original_stats['elapsed_time']
    time_percent = (time_diff / original_stats['elapsed_time']) * 100
    
    if time_diff > 0:
        print(f"   时间差异: +{time_diff:.2f}s ({time_percent:+.1f}%) - 模块化版本较慢")
    else:
        print(f"   时间差异: {time_diff:.2f}s ({time_percent:+.1f}%) - 模块化版本较快")
    
    print("\n📈 处理结果对比:")
    print(f"   文件总数: {original_stats['total_files']} vs {modular_stats['total_files']}")
    print(f"   成功处理: {original_stats['processed']} vs {modular_stats['processed']}")
    print(f"   失败文件: {original_stats['failed']} vs {modular_stats['failed']}")
    print(f"   成功率: {original_stats['success_rate']:.1f}% vs {modular_stats['success_rate']:.1f}%")
    
    # 计算改进
    success_diff = modular_stats['success_rate'] - original_stats['success_rate']
    processed_diff = modular_stats['processed'] - original_stats['processed']
    
    print("\n🎯 改进总结:")
    if success_diff > 0:
        print(f"   ✅ 成功率提升: +{success_diff:.1f}%")
    elif success_diff < 0:
        print(f"   ❌ 成功率下降: {success_diff:.1f}%")
    else:
        print(f"   ➖ 成功率持平")
    
    if processed_diff > 0:
        print(f"   ✅ 多处理文件: +{processed_diff} 个")
    elif processed_diff < 0:
        print(f"   ❌ 少处理文件: {processed_diff} 个")
    else:
        print(f"   ➖ 处理文件数持平")
    
    print("\n💡 总体评价:")
    if success_diff >= 0 and processed_diff >= 0:
        if time_percent < 20:  # 时间增加不超过20%
            print("   🎉 模块化版本整体表现优秀！")
        else:
            print("   ⚠️  模块化版本准确性更好，但性能有待优化")
    else:
        print("   ⚠️  模块化版本需要进一步优化")
    
    print("="*80)

if __name__ == "__main__":
    compare_batch_results()
