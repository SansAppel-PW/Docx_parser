#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用多种方法尝试将 EMF 流程图转换为 PNG 格式
"""

import os
import sys
import subprocess
from pathlib import Path

def convert_with_magick_command(emf_path, output_path):
    """使用 magick 命令行工具转换"""
    try:
        cmd = [
            'magick', 
            str(emf_path), 
            '-density', '150',
            '-background', 'white',
            '-flatten',
            str(output_path)
        ]
        
        print(f"执行命令：{' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                print(f"✅ magick 命令转换成功！")
                print(f"   输出文件：{output_path}")
                print(f"   文件大小：{file_size / 1024:.1f} KB")
                return True
        else:
            print(f"❌ magick 命令失败：{result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("❌ magick 命令超时")
        return False
    except Exception as e:
        print(f"❌ magick 命令执行错误：{e}")
        return False

def convert_with_sips(emf_path, output_path):
    """使用 macOS sips 工具转换"""
    try:
        cmd = ['sips', '-s', 'format', 'png', str(emf_path), '--out', str(output_path)]
        print(f"执行命令：{' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                print(f"✅ sips 转换成功！")
                print(f"   输出文件：{output_path}")
                print(f"   文件大小：{file_size / 1024:.1f} KB")
                return True
        else:
            print(f"❌ sips 失败：{result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ sips 执行错误：{e}")
        return False

def convert_with_convert_command(emf_path, output_path):
    """使用 convert 命令转换"""
    try:
        cmd = [
            'convert', 
            str(emf_path), 
            '-density', '150',
            '-background', 'white',
            '-flatten',
            str(output_path)
        ]
        
        print(f"执行命令：{' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                print(f"✅ convert 命令转换成功！")
                print(f"   输出文件：{output_path}")
                print(f"   文件大小：{file_size / 1024:.1f} KB")
                return True
        else:
            print(f"❌ convert 命令失败：{result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ convert 命令执行错误：{e}")
        return False

def main():
    """主函数"""
    print("=== EMF 到 PNG 转换工具 (多种方法) ===\n")
    
    # EMF 文件路径
    emf_file = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
    
    if not os.path.exists(emf_file):
        print(f"错误：找不到 EMF 文件：{emf_file}")
        return
    
    print(f"找到 EMF 文件：{emf_file}")
    
    # 获取文件信息
    file_size = os.path.getsize(emf_file)
    print(f"文件大小：{file_size / 1024:.1f} KB\n")
    
    # 输出文件路径
    base_dir = Path(emf_file).parent
    
    # 尝试不同的转换方法
    methods = [
        ('magick 命令', 'flowchart_magick.png', convert_with_magick_command),
        ('convert 命令', 'flowchart_convert.png', convert_with_convert_command),
        ('sips 工具', 'flowchart_sips.png', convert_with_sips),
    ]
    
    success_count = 0
    
    for method_name, output_filename, convert_func in methods:
        print(f"--- 尝试使用 {method_name} ---")
        output_path = base_dir / output_filename
        
        if convert_func(emf_file, output_path):
            success_count += 1
            print(f"✅ {method_name} 转换成功！\n")
        else:
            print(f"❌ {method_name} 转换失败\n")
    
    print(f"=== 转换完成：{success_count}/{len(methods)} 种方法成功 ===")
    
    if success_count > 0:
        print("\n✅ 至少有一种方法成功转换了流程图！")
        print("您可以在以下位置查看 PNG 文件：")
        for method_name, output_filename, _ in methods:
            output_path = base_dir / output_filename
            if os.path.exists(output_path):
                print(f"   - {output_path}")
    else:
        print("\n❌ 所有转换方法都失败了。")
        print("建议：")
        print("1. 使用在线转换工具：https://convertio.co/emf-png/")
        print("2. 使用原始的 EMF 文件（可用系统默认程序打开）")
        print("3. 使用我们之前生成的 HTML 预览文件")

if __name__ == "__main__":
    main()
