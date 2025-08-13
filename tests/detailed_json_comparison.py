#!/usr/bin/env python3
"""
详细对比两个版本的document.json内容
"""

import os
import sys
import json
import time
import subprocess
import logging
from difflib import SequenceMatcher

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def run_single_test(script_name, test_file, output_dir):
    """运行单个文件解析测试"""
    logger.info(f"🚀 运行 {script_name}...")
    
    # 创建临时脚本来修改输出目录
    temp_script = f"temp_{script_name}"
    
    try:
        # 读取原脚本
        with open(script_name, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 修改输出目录
        if script_name == "docx_parser.py":
            modified_content = content.replace(
                'output_dir = f"parsed_docs/single_files/{safe_name}"',
                f'output_dir = "{output_dir}"'
            )
        else:  # docx_parser_modular.py
            modified_content = content.replace(
                'output_dir = f"parsed_docs/single_files/{safe_name}"',
                f'output_dir = "{output_dir}"'
            )
        
        # 写入临时脚本
        with open(temp_script, 'w', encoding='utf-8') as f:
            f.write(modified_content)
        
        # 运行解析
        start_time = time.time()
        result = subprocess.run([
            sys.executable, temp_script, test_file
        ], capture_output=True, text=True, timeout=60)
        
        end_time = time.time()
        
        # 清理临时文件
        if os.path.exists(temp_script):
            os.remove(temp_script)
        
        if result.returncode == 0:
            logger.info(f"✅ {script_name} 运行成功，耗时 {end_time - start_time:.2f}s")
            return True
        else:
            logger.error(f"❌ {script_name} 运行失败:")
            logger.error(f"STDERR: {result.stderr}")
            return False
            
    except Exception as e:
        logger.error(f"❌ 运行 {script_name} 异常: {e}")
        # 清理临时文件
        if os.path.exists(temp_script):
            os.remove(temp_script)
        return False

def deep_compare_json(obj1, obj2, path=""):
    """深度对比两个JSON对象的差异"""
    differences = []
    
    if type(obj1) != type(obj2):
        differences.append({
            "path": path,
            "type": "type_mismatch",
            "original": str(type(obj1).__name__),
            "modular": str(type(obj2).__name__)
        })
        return differences
    
    if isinstance(obj1, dict):
        # 检查缺失的键
        for key in obj1:
            if key not in obj2:
                differences.append({
                    "path": f"{path}.{key}" if path else key,
                    "type": "missing_key_in_modular",
                    "original": f"存在 ({type(obj1[key]).__name__})",
                    "modular": "缺失"
                })
            else:
                differences.extend(deep_compare_json(obj1[key], obj2[key], f"{path}.{key}" if path else key))
        
        # 检查新增的键
        for key in obj2:
            if key not in obj1:
                differences.append({
                    "path": f"{path}.{key}" if path else key,
                    "type": "extra_key_in_modular",
                    "original": "缺失",
                    "modular": f"存在 ({type(obj2[key]).__name__})"
                })
    
    elif isinstance(obj1, list):
        if len(obj1) != len(obj2):
            differences.append({
                "path": f"{path}[length]",
                "type": "length_mismatch",
                "original": len(obj1),
                "modular": len(obj2)
            })
        
        min_len = min(len(obj1), len(obj2))
        for i in range(min_len):
            differences.extend(deep_compare_json(obj1[i], obj2[i], f"{path}[{i}]"))
    
    else:  # 基本类型
        if obj1 != obj2:
            # 对于字符串，计算相似度
            if isinstance(obj1, str) and isinstance(obj2, str):
                similarity = SequenceMatcher(None, obj1, obj2).ratio()
                differences.append({
                    "path": path,
                    "type": "value_mismatch",
                    "original": obj1[:100] + "..." if len(str(obj1)) > 100 else str(obj1),
                    "modular": obj2[:100] + "..." if len(str(obj2)) > 100 else str(obj2),
                    "similarity": f"{similarity:.2%}"
                })
            else:
                differences.append({
                    "path": path,
                    "type": "value_mismatch",
                    "original": obj1,
                    "modular": obj2
                })
    
    return differences

def analyze_json_content(original_path, modular_path):
    """分析两个JSON文件的内容差异"""
    logger.info("📄 分析JSON内容差异...")
    
    # 读取JSON文件
    try:
        with open(original_path, 'r', encoding='utf-8') as f:
            original_data = json.load(f)
    except Exception as e:
        logger.error(f"读取原始版本JSON失败: {e}")
        return None
    
    try:
        with open(modular_path, 'r', encoding='utf-8') as f:
            modular_data = json.load(f)
    except Exception as e:
        logger.error(f"读取模块化版本JSON失败: {e}")
        return None
    
    # 深度对比
    differences = deep_compare_json(original_data, modular_data)
    
    # 分析关键指标
    analysis = {
        "metadata_comparison": compare_metadata(
            original_data.get("metadata", {}),
            modular_data.get("metadata", {})
        ),
        "sections_comparison": compare_sections(
            original_data.get("sections", []),
            modular_data.get("sections", [])
        ),
        "images_comparison": compare_images(
            original_data.get("images", {}),
            modular_data.get("images", {})
        ),
        "processing_info_comparison": compare_processing_info(
            original_data.get("processing_info", {}),
            modular_data.get("processing_info", {})
        ),
        "detailed_differences": differences[:50]  # 限制显示前50个差异
    }
    
    return analysis

def compare_metadata(original_meta, modular_meta):
    """对比元数据"""
    comparison = {}
    
    key_fields = ["title", "author", "created", "modified", "file_size"]
    for field in key_fields:
        original_val = original_meta.get(field, "N/A")
        modular_val = modular_meta.get(field, "N/A")
        comparison[field] = {
            "original": original_val,
            "modular": modular_val,
            "match": original_val == modular_val
        }
    
    return comparison

def compare_sections(original_sections, modular_sections):
    """对比章节结构"""
    return {
        "count": {
            "original": len(original_sections),
            "modular": len(modular_sections),
            "match": len(original_sections) == len(modular_sections)
        },
        "structure_analysis": analyze_section_structure(original_sections, modular_sections)
    }

def analyze_section_structure(original_sections, modular_sections):
    """分析章节结构"""
    def count_content_types(sections):
        counts = {"paragraphs": 0, "tables": 0, "images": 0, "smartart": 0, "embedded_objects": 0}
        
        def traverse_section(section):
            content = section.get("content", [])
            for item in content:
                item_type = item.get("type", "unknown")
                if item_type == "paragraph":
                    counts["paragraphs"] += 1
                elif item_type == "table":
                    counts["tables"] += 1
                elif item_type == "image":
                    counts["images"] += 1
                elif item_type == "smartart":
                    counts["smartart"] += 1
                elif item_type == "embedded_object":
                    counts["embedded_objects"] += 1
                elif item_type == "section":
                    traverse_section(item)
        
        for section in sections:
            traverse_section(section)
        return counts
    
    original_counts = count_content_types(original_sections)
    modular_counts = count_content_types(modular_sections)
    
    return {
        "original": original_counts,
        "modular": modular_counts,
        "differences": {
            key: modular_counts[key] - original_counts[key]
            for key in original_counts
        }
    }

def compare_images(original_images, modular_images):
    """对比图片信息"""
    return {
        "count": {
            "original": len(original_images),
            "modular": len(modular_images),
            "difference": len(modular_images) - len(original_images)
        }
    }

def compare_processing_info(original_info, modular_info):
    """对比处理信息"""
    key_metrics = ["blocks_processed", "tables_found", "images_found"]
    comparison = {}
    
    for metric in key_metrics:
        original_val = original_info.get(metric, 0)
        modular_val = modular_info.get(metric, 0)
        comparison[metric] = {
            "original": original_val,
            "modular": modular_val,
            "difference": modular_val - original_val
        }
    
    return comparison

def print_detailed_comparison(analysis):
    """打印详细的对比结果"""
    print("\n" + "="*100)
    print("📊 DOCX Parser JSON内容详细对比报告")
    print("="*100)
    
    # 元数据对比
    print("\n📋 元数据对比:")
    metadata = analysis["metadata_comparison"]
    for field, data in metadata.items():
        match_icon = "✅" if data["match"] else "❌"
        print(f"   {match_icon} {field}: {data['original']} → {data['modular']}")
    
    # 章节结构对比
    print("\n📚 章节结构对比:")
    sections = analysis["sections_comparison"]
    count_match = "✅" if sections["count"]["match"] else "❌"
    print(f"   {count_match} 章节数量: {sections['count']['original']} → {sections['count']['modular']}")
    
    # 内容类型分析
    print("\n📝 内容类型分析:")
    structure = sections["structure_analysis"]
    for content_type, diff in structure["differences"].items():
        if diff != 0:
            icon = "⬆️" if diff > 0 else "⬇️"
            print(f"   {icon} {content_type}: {structure['original'][content_type]} → {structure['modular'][content_type]} (差异: {diff:+d})")
        else:
            print(f"   ✅ {content_type}: {structure['original'][content_type]} (一致)")
    
    # 图片对比
    print("\n🖼️  图片对比:")
    images = analysis["images_comparison"]
    diff = images["count"]["difference"]
    icon = "⬆️" if diff > 0 else "⬇️" if diff < 0 else "✅"
    print(f"   {icon} 图片数量: {images['count']['original']} → {images['count']['modular']} (差异: {diff:+d})")
    
    # 处理信息对比
    print("\n⚙️  处理信息对比:")
    processing = analysis["processing_info_comparison"]
    for metric, data in processing.items():
        diff = data["difference"]
        icon = "⬆️" if diff > 0 else "⬇️" if diff < 0 else "✅"
        print(f"   {icon} {metric}: {data['original']} → {data['modular']} (差异: {diff:+d})")
    
    # 详细差异
    differences = analysis["detailed_differences"]
    if differences:
        print(f"\n⚠️  发现 {len(differences)} 处差异 (显示前10个):")
        for i, diff in enumerate(differences[:10], 1):
            print(f"   {i:2d}. 路径: {diff['path']}")
            print(f"       类型: {diff['type']}")
            print(f"       原始: {diff['original']}")
            print(f"       模块: {diff['modular']}")
            if 'similarity' in diff:
                print(f"       相似度: {diff['similarity']}")
            print()
    else:
        print("\n✅ 未发现显著的内容差异")
    
    print("="*100)

def main():
    """主函数"""
    test_file = "Files/PLM2.0/BOM 审核申请.docx"
    
    if not os.path.exists(test_file):
        logger.error(f"测试文件不存在: {test_file}")
        return
    
    logger.info("🔍 开始详细JSON内容对比测试")
    logger.info(f"📄 测试文件: {test_file}")
    
    # 创建输出目录
    original_dir = "json_comparison/original_result"
    modular_dir = "json_comparison/modular_result"
    
    os.makedirs(original_dir, exist_ok=True)
    os.makedirs(modular_dir, exist_ok=True)
    
    # 运行两个版本
    logger.info("\n" + "="*60)
    success1 = run_single_test("docx_parser.py", test_file, original_dir)
    
    logger.info("\n" + "="*60)
    success2 = run_single_test("docx_parser_modular.py", test_file, modular_dir)
    
    if not success1 or not success2:
        logger.error("❌ 一个或多个测试失败")
        return
    
    # 对比JSON内容
    original_json = os.path.join(original_dir, "document.json")
    modular_json = os.path.join(modular_dir, "document.json")
    
    if not os.path.exists(original_json) or not os.path.exists(modular_json):
        logger.error("❌ JSON文件不存在")
        return
    
    analysis = analyze_json_content(original_json, modular_json)
    
    if analysis:
        print_detailed_comparison(analysis)
        
        # 保存详细报告
        report_file = "json_comparison/detailed_comparison_report.json"
        with open(report_file, "w", encoding="utf-8") as f:
            json.dump(analysis, f, ensure_ascii=False, indent=2)
        
        logger.info(f"\n📄 详细报告已保存到: {report_file}")
    
    logger.info("\n🎉 JSON内容对比完成!")

if __name__ == "__main__":
    main()
