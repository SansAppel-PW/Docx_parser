#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EMF 文件自动转换后处理器
在文档解析完成后，自动处理所有的 EMF 文件并尝试转换为 PNG
"""

import os
import json
import glob
import logging
from pathlib import Path

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def find_emf_files(output_dir):
    """
    在输出目录中查找所有 EMF 文件
    """
    emf_files = []
    
    # 查找所有 images 目录下的 EMF 文件
    image_dirs = glob.glob(os.path.join(output_dir, "**/images"), recursive=True)
    
    for img_dir in image_dirs:
        emf_patterns = [
            os.path.join(img_dir, "*.emf"),
            os.path.join(img_dir, "*.EMF")
        ]
        
        for pattern in emf_patterns:
            emf_files.extend(glob.glob(pattern))
    
    return emf_files

def convert_emf_with_python_tools(emf_path):
    """
    使用纯 Python 工具尝试转换 EMF
    """
    try:
        from PIL import Image
        import io
        
        # 输出 PNG 路径
        png_path = emf_path.replace('.emf', '.png').replace('.EMF', '.png')
        
        logger.info(f"尝试转换: {os.path.basename(emf_path)}")
        
        # 尝试用 PIL 读取
        with Image.open(emf_path) as img:
            logger.info(f"  检测到: {img.format}, 模式={img.mode}, 尺寸={img.size}")
            
            # 转换为 RGB 模式
            if img.mode in ('RGBA', 'LA'):
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'RGBA':
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # 保存为 PNG
            img.save(png_path, 'PNG', optimize=True)
            
            if os.path.exists(png_path) and os.path.getsize(png_path) > 0:
                logger.info(f"  ✅ 转换成功: {os.path.basename(png_path)}")
                
                # 创建转换报告
                report = {
                    "original": os.path.basename(emf_path),
                    "converted": os.path.basename(png_path),
                    "method": "PIL",
                    "status": "success",
                    "original_size_kb": f"{os.path.getsize(emf_path)/1024:.1f}",
                    "converted_size_kb": f"{os.path.getsize(png_path)/1024:.1f}"
                }
                
                return png_path, report
            else:
                logger.warning(f"  ❌ 转换失败: 输出文件无效")
                return None, None
                
    except ImportError:
        logger.warning("PIL 库未安装")
        return None, None
    except Exception as e:
        logger.warning(f"  ❌ PIL 转换失败: {e}")
        return None, None

def convert_emf_with_system_tools(emf_path):
    """
    使用系统工具尝试转换 EMF
    """
    import subprocess
    
    png_path = emf_path.replace('.emf', '.png').replace('.EMF', '.png')
    
    # macOS 系统工具
    system_commands = [
        # sips (macOS 内置)
        ['sips', '-s', 'format', 'png', emf_path, '--out', png_path],
        # Quick Look (macOS)
        ['qlmanage', '-t', '-s', '1024', '-o', os.path.dirname(png_path), emf_path]
    ]
    
    for cmd in system_commands:
        try:
            logger.info(f"  尝试命令: {' '.join(cmd[:2])}")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=20)
            
            # 检查可能的输出文件
            possible_outputs = [
                png_path,
                emf_path + '.png',
                os.path.join(os.path.dirname(png_path), os.path.basename(emf_path) + '.png')
            ]
            
            for output in possible_outputs:
                if os.path.exists(output) and os.path.getsize(output) > 0:
                    # 重命名到正确位置
                    if output != png_path:
                        os.rename(output, png_path)
                    
                    logger.info(f"  ✅ {cmd[0]} 转换成功")
                    
                    report = {
                        "original": os.path.basename(emf_path),
                        "converted": os.path.basename(png_path),
                        "method": cmd[0],
                        "status": "success",
                        "original_size_kb": f"{os.path.getsize(emf_path)/1024:.1f}",
                        "converted_size_kb": f"{os.path.getsize(png_path)/1024:.1f}"
                    }
                    
                    return png_path, report
                    
        except subprocess.TimeoutExpired:
            logger.warning(f"  ⏰ {cmd[0]} 超时")
        except FileNotFoundError:
            logger.debug(f"  ❓ {cmd[0]} 未找到")
        except Exception as e:
            logger.warning(f"  ❌ {cmd[0]} 失败: {e}")
    
    return None, None

def update_document_json(output_dir, conversion_reports):
    """
    更新 document.json 文件，将 EMF 路径替换为 PNG 路径
    """
    updated_files = []
    
    # 查找所有 document.json 文件
    json_files = glob.glob(os.path.join(output_dir, "**/document.json"), recursive=True)
    
    for json_file in json_files:
        try:
            # 读取 JSON
            with open(json_file, 'r', encoding='utf-8') as f:
                doc_data = json.load(f)
            
            # 更新成功转换的图像引用
            updated = False
            
            def update_content(obj):
                nonlocal updated
                if isinstance(obj, dict):
                    # 更新图像路径
                    if obj.get('type') == 'embedded_object' and 'preview_image' in obj:
                        old_path = obj['preview_image']
                        if old_path.endswith('.emf'):
                            new_path = old_path.replace('.emf', '.png')
                            # 检查是否有对应的转换报告
                            emf_name = os.path.basename(old_path)
                            for report in conversion_reports:
                                if report['original'] == emf_name:
                                    obj['preview_image'] = new_path
                                    obj['format'] = 'png'  # 更新格式
                                    obj['conversion_method'] = report['method']
                                    updated = True
                                    logger.info(f"    更新引用: {emf_name} -> {report['converted']}")
                                    break
                    
                    # 递归处理嵌套对象
                    for value in obj.values():
                        update_content(value)
                elif isinstance(obj, list):
                    for item in obj:
                        update_content(item)
            
            update_content(doc_data)
            
            # 如果有更新，保存文件
            if updated:
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(doc_data, f, ensure_ascii=False, indent=2)
                updated_files.append(json_file)
                logger.info(f"  📝 更新了: {os.path.relpath(json_file, output_dir)}")
        
        except Exception as e:
            logger.error(f"更新 JSON 文件失败 {json_file}: {e}")
    
    return updated_files

def create_conversion_summary(output_dir, conversion_reports):
    """
    创建转换总结报告
    """
    try:
        summary = {
            "conversion_timestamp": str(Path().resolve()),
            "total_emf_files": len([r for r in conversion_reports if r is not None]),
            "successful_conversions": len([r for r in conversion_reports if r and r['status'] == 'success']),
            "failed_conversions": len([r for r in conversion_reports if r and r['status'] == 'failed']),
            "conversion_details": [r for r in conversion_reports if r is not None]
        }
        
        summary_path = os.path.join(output_dir, "emf_conversion_summary.json")
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        logger.info(f"📊 转换总结保存到: {summary_path}")
        return summary_path
        
    except Exception as e:
        logger.error(f"创建转换总结失败: {e}")
        return None

def main(output_dir):
    """
    主函数：处理指定目录中的所有 EMF 文件
    """
    logger.info(f"🔍 开始处理目录: {output_dir}")
    
    # 查找所有 EMF 文件
    emf_files = find_emf_files(output_dir)
    
    if not emf_files:
        logger.info("✅ 没有找到需要转换的 EMF 文件")
        return
    
    logger.info(f"📁 找到 {len(emf_files)} 个 EMF 文件")
    
    conversion_reports = []
    successful_conversions = 0
    
    # 逐个转换 EMF 文件
    for emf_file in emf_files:
        logger.info(f"\n🔄 处理: {os.path.basename(emf_file)}")
        
        # 先尝试 Python 工具
        result, report = convert_emf_with_python_tools(emf_file)
        
        # 如果失败，尝试系统工具
        if not result:
            result, report = convert_emf_with_system_tools(emf_file)
        
        # 记录结果
        if result and report:
            successful_conversions += 1
            conversion_reports.append(report)
        else:
            # 记录失败
            failed_report = {
                "original": os.path.basename(emf_file),
                "converted": None,
                "method": "none",
                "status": "failed",
                "original_size_kb": f"{os.path.getsize(emf_file)/1024:.1f}",
                "note": "所有转换方法都失败，建议使用在线转换工具"
            }
            conversion_reports.append(failed_report)
    
    # 更新 document.json 文件
    logger.info(f"\n📝 更新文档引用...")
    updated_files = update_document_json(output_dir, 
                                       [r for r in conversion_reports if r['status'] == 'success'])
    
    # 创建转换总结
    summary_path = create_conversion_summary(output_dir, conversion_reports)
    
    # 输出最终报告
    logger.info(f"\n" + "="*60)
    logger.info(f"🎉 EMF 转换处理完成!")
    logger.info(f"="*60)
    logger.info(f"📊 总计 EMF 文件: {len(emf_files)}")
    logger.info(f"✅ 成功转换: {successful_conversions}")
    logger.info(f"❌ 转换失败: {len(emf_files) - successful_conversions}")
    logger.info(f"📝 更新的文档: {len(updated_files)}")
    
    if summary_path:
        logger.info(f"📋 详细报告: {os.path.relpath(summary_path, output_dir)}")
    
    if successful_conversions > 0:
        logger.info(f"\n🎯 现在您的文档中的流程图已自动转换为 PNG 格式！")
    
    if len(emf_files) - successful_conversions > 0:
        logger.info(f"\n💡 对于转换失败的文件，建议使用在线工具：")
        logger.info(f"   https://convertio.co/emf-png/")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        output_dir = sys.argv[1]
    else:
        output_dir = "Files/test_embedded_objects"  # 默认目录
    
    if not os.path.exists(output_dir):
        print(f"错误: 目录不存在 {output_dir}")
        sys.exit(1)
    
    main(output_dir)
