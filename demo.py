#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX 解析器演示脚本

这个脚本演示了如何使用 DOCX 解析器来处理示例文档，
展示解析器的主要功能和输出格式。

作者: DOCX Parser Team
版本: 1.0.0
"""

import os
import sys
import time
import json
from pathlib import Path

# 添加 src 目录到 Python 路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

try:
    from src.parsers.document_parser import parse_docx
except ImportError as e:
    print(f"❌ 导入错误: {e}")
    print("请确保您在项目根目录下运行此脚本")
    sys.exit(1)


def print_banner():
    """打印项目横幅"""
    banner = """
╔══════════════════════════════════════════════════════════════╗
║                    DOCX 解析器演示程序                        ║
║                                                              ║
║  🚀 功能强大的 Word 文档解析工具                              ║
║  📝 支持文本、图片、SmartArt、嵌入对象等内容提取               ║
║  ⚡ 模块化架构，高性能处理                                   ║
╚══════════════════════════════════════════════════════════════╝
"""
    print(banner)


def print_separator(title=""):
    """打印分隔线"""
    if title:
        print(f"\n{'='*20} {title} {'='*20}")
    else:
        print("\n" + "="*60)


def check_example_file():
    """检查示例文件是否存在"""
    example_file = Path("Files/examples/demo.docx")
    
    if not example_file.exists():
        print("⚠️  未找到示例文档 'Files/examples/demo.docx'")
        print("\n请确保示例文件存在，或者将您的 DOCX 文件放置在以下位置：")
        print(f"   📁 {example_file.absolute()}")
        
        # 检查是否有其他 DOCX 文件
        examples_dir = Path("Files/examples")
        if examples_dir.exists():
            docx_files = list(examples_dir.glob("*.docx"))
            if docx_files:
                print(f"\n💡 发现其他 DOCX 文件:")
                for i, file in enumerate(docx_files, 1):
                    print(f"   {i}. {file.name}")
                return docx_files[0]  # 返回第一个找到的文件
        
        return None
    
    return example_file


def display_results_summary(result, processing_time):
    """显示解析结果摘要"""
    print_separator("解析结果摘要")
    
    # 基本信息
    metadata = result.get('metadata', {})
    print(f"📄 文档标题: {metadata.get('title', '未知')}")
    print(f"👤 作者: {metadata.get('author', '未知')}")
    print(f"⏱️  处理时间: {processing_time:.2f} 秒")
    
    # 内容统计
    sections = result.get('sections', [])
    images = result.get('images', {})
    smartart = result.get('smartart', {})
    
    print(f"\n📊 内容统计:")
    print(f"   📝 段落数量: {len([s for s in sections if s.get('type') == 'paragraph'])}")
    print(f"   📋 表格数量: {len([s for s in sections if s.get('type') == 'table'])}")
    print(f"   🖼️  图片数量: {len(images)}")
    print(f"   🎨 SmartArt: {len(smartart)}")
    
    # 处理信息
    processing_info = result.get('processing_info', {})
    warnings = processing_info.get('warnings', [])
    errors = processing_info.get('errors', [])
    
    if warnings:
        print(f"\n⚠️  警告 ({len(warnings)} 个):")
        for warning in warnings[:3]:  # 只显示前3个
            print(f"   • {warning}")
        if len(warnings) > 3:
            print(f"   • ... 还有 {len(warnings) - 3} 个警告")
    
    if errors:
        print(f"\n❌ 错误 ({len(errors)} 个):")
        for error in errors[:3]:  # 只显示前3个
            print(f"   • {error}")
        if len(errors) > 3:
            print(f"   • ... 还有 {len(errors) - 3} 个错误")


def display_detailed_results(result):
    """显示详细解析结果"""
    print_separator("详细解析结果")
    
    # 显示前几个段落
    sections = result.get('sections', [])
    paragraphs = [s for s in sections if s.get('type') == 'paragraph']
    
    if paragraphs:
        print("📝 文档段落预览 (前3段):")
        for i, para in enumerate(paragraphs[:3], 1):
            text = para.get('text', '').strip()
            if text:
                preview = text[:100] + "..." if len(text) > 100 else text
                print(f"   {i}. {preview}")
    
    # 显示图片信息
    images = result.get('images', {})
    if images:
        print(f"\n🖼️  图片详细信息:")
        for img_id, img_info in list(images.items())[:3]:  # 只显示前3个
            print(f"   📄 {img_id}:")
            print(f"      格式: {img_info.get('format', '未知')}")
            print(f"      尺寸: {img_info.get('width', 0)} x {img_info.get('height', 0)}")
            print(f"      大小: {img_info.get('size', '未知')}")
        
        if len(images) > 3:
            print(f"   ... 还有 {len(images) - 3} 个图片")
    
    # 显示 SmartArt 信息
    smartart = result.get('smartart', {})
    if smartart:
        print(f"\n🎨 SmartArt 详细信息:")
        for art_id, art_info in list(smartart.items())[:2]:  # 只显示前2个
            print(f"   🎯 {art_id}:")
            print(f"      标题: {art_info.get('title', '未知')}")
            print(f"      类型: {art_info.get('type', '未知')}")
            nodes = art_info.get('nodes', [])
            print(f"      节点数: {len(nodes)}")


def save_results_preview(result, output_dir):
    """保存结果预览文件"""
    try:
        # 创建简化的结果预览
        preview = {
            "文档信息": result.get('metadata', {}),
            "内容统计": {
                "段落数": len([s for s in result.get('sections', []) if s.get('type') == 'paragraph']),
                "表格数": len([s for s in result.get('sections', []) if s.get('type') == 'table']),
                "图片数": len(result.get('images', {})),
                "SmartArt数": len(result.get('smartart', {}))
            },
            "处理信息": result.get('processing_info', {}),
            "注意": "这是演示程序生成的结果预览，完整结果请查看详细的JSON文件"
        }
        
        preview_file = os.path.join(output_dir, "demo_result_preview.json")
        with open(preview_file, 'w', encoding='utf-8') as f:
            json.dump(preview, f, ensure_ascii=False, indent=2)
        
        print(f"\n💾 结果预览已保存: {preview_file}")
        
    except Exception as e:
        print(f"⚠️  保存预览文件时出错: {e}")


def main():
    """主函数"""
    print_banner()
    
    # 检查示例文件
    print("🔍 检查示例文档...")
    example_file = check_example_file()
    
    if not example_file:
        print("\n❌ 无法找到可用的示例文档")
        print("\n💡 解决方案:")
        print("   1. 将您的 DOCX 文件重命名为 'demo.docx' 并放置在 'Files/examples/' 目录下")
        print("   2. 或者修改此脚本指向您的文档文件")
        return
    
    print(f"✅ 找到示例文档: {example_file.name}")
    
    # 设置输出目录
    output_dir = "parsed_docs/examples"
    os.makedirs(output_dir, exist_ok=True)
    
    print_separator("开始解析")
    print(f"📁 输入文件: {example_file}")
    print(f"📁 输出目录: {output_dir}")
    print(f"⚙️  处理模式: 标准模式")
    
    try:
        # 开始解析
        print("\n🚀 开始解析文档...")
        start_time = time.time()
        
        result = parse_docx(str(example_file), output_dir)
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        print(f"✅ 解析完成! 耗时 {processing_time:.2f} 秒")
        
        # 显示结果
        display_results_summary(result, processing_time)
        display_detailed_results(result)
        
        # 保存预览
        save_results_preview(result, output_dir)
        
        print_separator("使用建议")
        print("💡 下一步您可以:")
        print("   1. 查看输出目录中的详细解析结果")
        print("   2. 使用 'python docx_parser_modular.py' 处理其他文档")
        print("   3. 参考 README.md 了解更多高级功能")
        print("   4. 尝试批量处理功能")
        
        print_separator()
        print("🎉 演示完成! 感谢使用 DOCX 解析器!")
        
    except Exception as e:
        print(f"\n❌ 解析过程中出现错误:")
        print(f"   错误类型: {type(e).__name__}")
        print(f"   错误详情: {str(e)}")
        print(f"\n🔧 故障排除建议:")
        print(f"   1. 确保文档文件完整且未损坏")
        print(f"   2. 检查是否安装了所有必需的依赖")
        print(f"   3. 尝试使用快速模式: quick_mode=True")
        print(f"   4. 查看日志文件获取更多信息")
        
        return False
    
    return True


if __name__ == "__main__":
    # 运行演示
    success = main()
    
    # 退出代码
    sys.exit(0 if success else 1)
