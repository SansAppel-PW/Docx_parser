#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用 Wand (ImageMagick Python 绑定) 将 EMF 流程图转换为 PNG 格式
"""

import os
import sys
from pathlib import Path

# 设置 ImageMagick 环境变量
os.environ['MAGICK_HOME'] = '/opt/homebrew'
if 'DYLD_LIBRARY_PATH' in os.environ:
    os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + os.environ['DYLD_LIBRARY_PATH']
else:
    os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib'

try:
    from wand.image import Image
    from wand.exceptions import WandException
except ImportError:
    print("错误：未找到 Wand 库，请安装：pip install Wand")
    sys.exit(1)

def convert_emf_to_png_wand(emf_path, output_path=None, resolution=150):
    """
    使用 Wand 将 EMF 文件转换为 PNG
    
    Args:
        emf_path: EMF 文件路径
        output_path: 输出 PNG 文件路径（可选）
        resolution: 分辨率 DPI
    
    Returns:
        成功返回输出文件路径，失败返回 None
    """
    try:
        # 检查输入文件
        if not os.path.exists(emf_path):
            print(f"错误：EMF 文件不存在：{emf_path}")
            return None
        
        # 设置输出路径
        if output_path is None:
            base_name = Path(emf_path).stem
            output_dir = Path(emf_path).parent
            output_path = output_dir / f"{base_name}.png"
        
        print(f"开始转换：{emf_path}")
        print(f"输出路径：{output_path}")
        print(f"分辨率：{resolution} DPI")
        
        # 使用 Wand 进行转换
        with Image() as img:
            # 设置分辨率
            img.resolution = (resolution, resolution)
            
            # 读取 EMF 文件
            img.read(filename=str(emf_path))
            
            # 设置背景色为白色
            img.background_color = 'white'
            img.alpha_channel = 'remove'
            
            # 设置格式为 PNG
            img.format = 'png'
            
            # 保存为 PNG
            img.save(filename=str(output_path))
        
        # 检查输出文件
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"✅ 转换成功！")
            print(f"   输出文件：{output_path}")
            print(f"   文件大小：{file_size / 1024:.1f} KB")
            return str(output_path)
        else:
            print("❌ 转换失败：输出文件未生成")
            return None
            
    except WandException as e:
        print(f"❌ Wand 转换错误：{e}")
        return None
    except Exception as e:
        print(f"❌ 转换过程中发生错误：{e}")
        return None

def main():
    """主函数"""
    print("=== EMF 到 PNG 转换工具 (使用 Wand) ===\n")
    
    # EMF 文件路径
    emf_file = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
    
    if not os.path.exists(emf_file):
        print(f"错误：找不到 EMF 文件：{emf_file}")
        return
    
    print(f"找到 EMF 文件：{emf_file}")
    
    # 获取文件信息
    file_size = os.path.getsize(emf_file)
    print(f"文件大小：{file_size / 1024:.1f} KB")
    
    # 尝试不同分辨率的转换
    resolutions = [150, 300, 200]
    
    for resolution in resolutions:
        print(f"\n--- 尝试 {resolution} DPI 转换 ---")
        
        # 设置输出文件名
        output_file = f"Files/test_embedded_objects/images/flowchart_{resolution}dpi.png"
        
        # 转换
        result = convert_emf_to_png_wand(emf_file, output_file, resolution)
        
        if result:
            print(f"✅ {resolution} DPI 转换成功！")
            break
        else:
            print(f"❌ {resolution} DPI 转换失败")
    
    print("\n=== 转换完成 ===")

if __name__ == "__main__":
    main()
