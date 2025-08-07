#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将EMF格式的图像文件转换为PNG格式的工具
"""

import os
import sys
from PIL import Image
import subprocess

def convert_emf_to_png_with_wand():
    """使用Wand库（ImageMagick的Python绑定）转换EMF到PNG"""
    try:
        from wand.image import Image as WandImage
        
        emf_path = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
        png_path = "Files/test_embedded_objects/images/embedded_preview_converted.png"
        
        with WandImage(filename=emf_path) as img:
            # 转换为PNG格式
            img.format = 'png'
            # 设置背景色为白色
            img.background_color = 'white'
            # 合并图层
            img = img.merge_layers('flatten')
            # 保存
            img.save(filename=png_path)
            
        print(f"成功使用Wand转换: {png_path}")
        return png_path
        
    except ImportError:
        print("Wand库未安装，请运行: pip install Wand")
        return None
    except Exception as e:
        print(f"Wand转换失败: {e}")
        return None

def convert_emf_with_system_tools():
    """尝试使用系统工具转换EMF"""
    emf_path = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
    
    # 方法1: 尝试ImageMagick的convert命令
    png_path1 = "Files/test_embedded_objects/images/embedded_preview_imagemagick.png"
    try:
        result = subprocess.run([
            'magick', emf_path, png_path1
        ], capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0 and os.path.exists(png_path1):
            print(f"成功使用ImageMagick转换: {png_path1}")
            return png_path1
        else:
            print(f"ImageMagick转换失败: {result.stderr}")
    except Exception as e:
        print(f"ImageMagick转换错误: {e}")
    
    # 方法2: 尝试convert命令（旧版本ImageMagick）
    png_path2 = "Files/test_embedded_objects/images/embedded_preview_convert.png"
    try:
        result = subprocess.run([
            'convert', emf_path, png_path2
        ], capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0 and os.path.exists(png_path2):
            print(f"成功使用convert转换: {png_path2}")
            return png_path2
        else:
            print(f"convert转换失败: {result.stderr}")
    except Exception as e:
        print(f"convert转换错误: {e}")
    
    return None

def analyze_emf_file():
    """分析EMF文件的信息"""
    emf_path = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
    
    if not os.path.exists(emf_path):
        print(f"EMF文件不存在: {emf_path}")
        return
    
    file_size = os.path.getsize(emf_path)
    print(f"EMF文件大小: {file_size} 字节 ({file_size/1024:.2f} KB)")
    
    # 读取文件头信息
    with open(emf_path, 'rb') as f:
        header = f.read(88)  # EMF头部大小
        if len(header) >= 88:
            # EMF签名应该是 0x00000001
            signature = int.from_bytes(header[40:44], 'little')
            print(f"EMF签名: 0x{signature:08X}")
            
            # 图像边界
            left = int.from_bytes(header[8:12], 'little')
            top = int.from_bytes(header[12:16], 'little') 
            right = int.from_bytes(header[16:20], 'little')
            bottom = int.from_bytes(header[20:24], 'little')
            
            width = right - left
            height = bottom - top
            print(f"图像尺寸: {width} x {height}")

def try_pillow_conversion():
    """尝试使用Pillow的不同方法"""
    emf_path = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
    
    print("尝试Pillow转换...")
    
    # 方法1: 直接打开
    try:
        with Image.open(emf_path) as img:
            png_path = "Files/test_embedded_objects/images/embedded_preview_pillow_direct.png"
            img.save(png_path, 'PNG')
            print(f"Pillow直接转换成功: {png_path}")
            return png_path
    except Exception as e:
        print(f"Pillow直接转换失败: {e}")
    
    return None

def install_dependencies():
    """尝试安装必要的依赖"""
    print("尝试安装图像处理依赖...")
    
    try:
        # 安装Wand
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'Wand'], 
                      capture_output=True, text=True)
        print("Wand安装完成")
    except Exception as e:
        print(f"Wand安装失败: {e}")
    
    try:
        # 安装Pillow（如果没有）
        subprocess.run([sys.executable, '-m', 'pip', 'install', '--upgrade', 'Pillow'], 
                      capture_output=True, text=True)
        print("Pillow安装/升级完成")
    except Exception as e:
        print(f"Pillow安装失败: {e}")

def main():
    print("EMF到PNG转换工具")
    print("=" * 40)
    
    # 分析EMF文件
    analyze_emf_file()
    print()
    
    # 尝试各种转换方法
    success = False
    
    # 方法1: Pillow
    result = try_pillow_conversion()
    if result:
        success = True
    
    # 方法2: 系统工具
    if not success:
        result = convert_emf_with_system_tools()
        if result:
            success = True
    
    # 方法3: Wand
    if not success:
        result = convert_emf_to_png_with_wand()
        if result:
            success = True
    
    if not success:
        print("\n所有转换方法都失败了。")
        print("建议:")
        print("1. 安装ImageMagick: brew install imagemagick (macOS)")
        print("2. 安装Wand: pip install Wand")
        print("3. 或者手动在其他软件中打开EMF文件并另存为PNG")
        
        # 尝试安装依赖
        install_choice = input("\n是否尝试自动安装依赖? (y/n): ")
        if install_choice.lower() == 'y':
            install_dependencies()
            
            # 重试Wand转换
            print("\n重试Wand转换...")
            result = convert_emf_to_png_with_wand()
            if result:
                success = True
    
    if success:
        print(f"\n✅ 转换成功!")
        print("您现在可以查看PNG格式的流程图图像。")
    else:
        print(f"\n❌ 转换失败")
        print("EMF文件已保存，您可以使用其他工具手动转换。")

if __name__ == "__main__":
    main()
