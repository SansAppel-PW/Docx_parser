#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用 PIL/Pillow 尝试将 EMF 转换为 PNG
"""

import os
import sys
from pathlib import Path

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    PILImage = None

def convert_with_pillow(emf_path, output_path):
    """使用 PIL/Pillow 转换"""
    if not PIL_AVAILABLE or PILImage is None:
        print("❌ PIL/Pillow 库未安装")
        return False
    
    try:
        print(f"尝试使用 PIL 打开：{emf_path}")
        
        # 尝试直接打开 EMF 文件
        with PILImage.open(emf_path) as img:
            print(f"✅ 成功打开 EMF 文件")
            print(f"   图像模式：{img.mode}")
            print(f"   图像尺寸：{img.size}")
            print(f"   图像格式：{img.format}")
            
            # 转换为 RGB 模式（如果需要）
            if img.mode == 'RGBA':
                # 创建白色背景
                background = PILImage.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[-1] if len(img.split()) == 4 else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # 保存为 PNG
            img.save(output_path, 'PNG', optimize=True)
            
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                print(f"✅ PIL 转换成功！")
                print(f"   输出文件：{output_path}")
                print(f"   文件大小：{file_size / 1024:.1f} KB")
                return True
            else:
                print("❌ 输出文件未生成")
                return False
                
    except Exception as e:
        print(f"❌ PIL 转换失败：{e}")
        return False

def create_info_file(emf_path):
    """创建 EMF 文件信息文件"""
    try:
        info_content = f"""# EMF 流程图文件信息

## 文件详情
- **文件名称**: {Path(emf_path).name}
- **文件路径**: {emf_path}
- **文件大小**: {os.path.getsize(emf_path) / 1024:.1f} KB
- **文件格式**: Enhanced Metafile Format (EMF)

## 内容描述
这是从 Word 文档中提取的 Visio 流程图的预览图像。该流程图描述了 BOM 审核申请流程中的详细步骤。

## 查看方法

### 方法 1: 在线转换
访问在线转换工具将 EMF 转换为 PNG：
- https://convertio.co/emf-png/
- https://www.freeconvert.com/emf-to-png

### 方法 2: 系统默认程序
在 macOS 上，您可以：
1. 双击 EMF 文件尝试用默认程序打开
2. 右键点击 → "打开方式" → 选择图像查看器

### 方法 3: 使用预览应用
如果系统支持，可以尝试用 "预览" 应用打开。

### 方法 4: HTML 预览
我们已经生成了 HTML 预览文件：
- `emf_preview.html` - 包含 Base64 编码的图像预览

## 技术说明
EMF 格式是 Windows 增强型图元文件格式，主要用于矢量图形。某些系统和工具可能不支持直接查看，但内容完整保存。

## 原始来源
此图像是从以下位置提取：
- Word 文档：BOM 审核申请.docx
- 位置：两段文字之间的嵌入式 Visio 对象
- 对象类型：Visio.Drawing.15
- 尺寸：1552x904 像素
"""
        
        info_file = Path(emf_path).parent / "flowchart_info.md"
        with open(info_file, 'w', encoding='utf-8') as f:
            f.write(info_content)
        
        print(f"✅ 已创建信息文件：{info_file}")
        return True
        
    except Exception as e:
        print(f"❌ 创建信息文件失败：{e}")
        return False

def main():
    """主函数"""
    print("=== EMF 到 PNG 转换工具 (PIL 方法) ===\n")
    
    # EMF 文件路径
    emf_file = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
    
    if not os.path.exists(emf_file):
        print(f"错误：找不到 EMF 文件：{emf_file}")
        return
    
    print(f"找到 EMF 文件：{emf_file}")
    
    # 获取文件信息
    file_size = os.path.getsize(emf_file)
    print(f"文件大小：{file_size / 1024:.1f} KB\n")
    
    # 输出路径
    output_file = "Files/test_embedded_objects/images/flowchart_pillow.png"
    
    print("--- 尝试使用 PIL/Pillow 转换 ---")
    success = convert_with_pillow(emf_file, output_file)
    
    print("\n--- 创建文件信息 ---")
    create_info_file(emf_file)
    
    print("\n=== 总结 ===")
    if success:
        print("✅ 转换成功！您现在有了 PNG 格式的流程图。")
    else:
        print("❌ PIL 转换失败，但我们已经为您提供了其他查看方法。")
        print("\n📖 请查看生成的信息文件了解如何查看流程图：")
        print("   Files/test_embedded_objects/images/flowchart_info.md")
        
        print("\n🌐 最简单的方法是使用在线转换：")
        print("   1. 访问：https://convertio.co/emf-png/")
        print("   2. 上传文件：embedded_preview_embedded_obj_ef975fd7.emf")
        print("   3. 下载转换后的 PNG 文件")

if __name__ == "__main__":
    main()
