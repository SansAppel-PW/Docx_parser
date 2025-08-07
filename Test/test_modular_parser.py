#!/usr/bin/env python3
"""
快速测试模块化DOCX解析器
"""

import os
import sys
import time
import logging

# 添加父目录到Python路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

def test_modular_parser():
    """测试模块化解析器"""
    print("🚀 测试模块化DOCX解析器")
    print("="*50)
    
    # 配置简单日志
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    
    # 测试文件
    test_file = "../Files/PLM2.0/BOM 审核申请.docx"
    
    if not os.path.exists(test_file):
        # 尝试不同的路径
        test_file = "Files/PLM2.0/BOM 审核申请.docx"
        
    if not os.path.exists(test_file):
        print("❌ 测试文件不存在，请确保有示例文档")
        print("💡 请检查 Files/PLM2.0/ 目录中是否有文档")
        return
    
    print(f"📄 测试文件: {os.path.basename(test_file)}")
    
    # 使用模块化解析器
    try:
        from docx_parser_modular import main as modular_main
        
        start_time = time.time()
        
        # 模拟命令行参数
        original_argv = sys.argv
        sys.argv = ["docx_parser_modular.py", test_file]
        
        try:
            modular_main()
            elapsed_time = time.time() - start_time
            print(f"✅ 模块化解析器测试成功")
            print(f"⏱️  处理时间: {elapsed_time:.2f} 秒")
        finally:
            sys.argv = original_argv
            
    except Exception as e:
        print(f"❌ 模块化解析器测试失败: {e}")

def test_batch_processing():
    """测试批量处理"""
    print("\n📁 测试批量处理")
    print("="*50)
    
    # 测试批量处理
    try:
        from docx_parser_modular import main as modular_main
        
        start_time = time.time()
        
        # 模拟命令行参数
        original_argv = sys.argv
        test_dir = "../Files/PLM2.0"
        
        if not os.path.exists(test_dir):
            test_dir = "Files/PLM2.0"
            
        sys.argv = ["docx_parser_modular.py", test_dir]
        
        try:
            modular_main()
            elapsed_time = time.time() - start_time
            print(f"✅ 批量处理测试成功")
            print(f"⏱️  总处理时间: {elapsed_time:.2f} 秒")
        finally:
            sys.argv = original_argv
            
    except Exception as e:
        print(f"❌ 批量处理测试失败: {e}")

if __name__ == "__main__":
    test_modular_parser()
    test_batch_processing()
    
    print("\n🎉 测试完成!")
    print("📁 请查看 parsed_docs/ 目录中的解析结果")
