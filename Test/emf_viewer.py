#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单的EMF文件信息查看器和Base64转换工具
"""

import os
import base64

def emf_to_base64():
    """将EMF文件转换为Base64编码，可以在HTML中显示"""
    emf_path = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
    
    if not os.path.exists(emf_path):
        print(f"EMF文件不存在: {emf_path}")
        return
    
    with open(emf_path, 'rb') as f:
        emf_data = f.read()
    
    # 转换为Base64
    base64_data = base64.b64encode(emf_data).decode('utf-8')
    
    # 创建HTML文件用于查看
    html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Visio流程图预览</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        .image-container {{
            text-align: center;
            margin: 20px 0;
            padding: 20px;
            border: 2px dashed #ccc;
            border-radius: 8px;
        }}
        .info {{
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Visio流程图预览</h1>
        
        <div class="info">
            <h3>文件信息:</h3>
            <ul>
                <li><strong>文件类型:</strong> EMF (Enhanced Metafile)</li>
                <li><strong>文件大小:</strong> {len(emf_data)/1024:.2f} KB</li>
                <li><strong>原始路径:</strong> {emf_path}</li>
                <li><strong>来源:</strong> 嵌入的Visio绘图对象</li>
            </ul>
        </div>
        
        <div class="image-container">
            <h3>图像预览:</h3>
            <p style="color: #666;">
                注意: EMF格式在不同浏览器中的支持可能有限。<br>
                如果图像不显示，请尝试以下方法：
            </p>
            
            <!-- 尝试直接显示EMF -->
            <img src="data:image/emf;base64,{base64_data}" 
                 alt="Visio流程图" 
                 style="max-width: 100%; border: 1px solid #ddd; margin: 10px 0;"/>
            
            <!-- 也尝试作为通用二进制文件 -->
            <br>
            <img src="data:application/octet-stream;base64,{base64_data}" 
                 alt="Visio流程图 (通用格式)" 
                 style="max-width: 100%; border: 1px solid #ddd; margin: 10px 0;"/>
        </div>
        
        <div class="info">
            <h3>📋 建议的查看方法:</h3>
            <ol>
                <li><strong>Windows系统:</strong> 可以直接双击EMF文件用默认程序打开</li>
                <li><strong>Mac系统:</strong> 可以用Preview应用打开EMF文件</li>
                <li><strong>在线转换:</strong> 
                    <ul>
                        <li>访问 <a href="https://convertio.co/emf-png/" target="_blank">Convertio EMF转PNG</a></li>
                        <li>或 <a href="https://www.freeconvert.com/emf-to-png" target="_blank">FreeConvert EMF转PNG</a></li>
                    </ul>
                </li>
                <li><strong>下载文件:</strong> 
                    <a href="data:application/octet-stream;base64,{base64_data}" 
                       download="visio_flowchart.emf">点击下载EMF文件</a>
                </li>
            </ol>
        </div>
        
        <div class="info">
            <h3>🔍 Base64数据 (前200字符):</h3>
            <code style="word-break: break-all; background: #f1f1f1; padding: 10px; display: block; border-radius: 3px;">
                {base64_data[:200]}...
            </code>
            <p><small>完整长度: {len(base64_data)} 字符</small></p>
        </div>
    </div>
</body>
</html>
"""
    
    html_path = "Files/test_embedded_objects/visio_flowchart_preview.html"
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✅ 已创建HTML预览文件: {html_path}")
    print(f"📁 EMF文件大小: {len(emf_data)/1024:.2f} KB")
    print(f"📊 Base64编码长度: {len(base64_data)} 字符")
    print("\n🌐 请用浏览器打开HTML文件查看流程图")
    
    return html_path

def create_simple_converter():
    """创建一个简单的Python转换脚本供用户使用"""
    converter_script = '''#!/usr/bin/env python3
"""
简单的EMF文件转换工具
使用系统的转换命令将EMF转换为PNG
"""

import os
import subprocess
import sys

def convert_emf():
    emf_file = "embedded_preview_embedded_obj_ef975fd7.emf"
    png_file = "visio_flowchart.png"
    
    if not os.path.exists(emf_file):
        print(f"❌ EMF文件不存在: {emf_file}")
        return False
    
    # 尝试不同的转换命令
    commands = [
        ['magick', emf_file, png_file],
        ['convert', emf_file, png_file],
        ['sips', '-s', 'format', 'png', emf_file, '--out', png_file],  # macOS
    ]
    
    for cmd in commands:
        try:
            print(f"🔄 尝试命令: {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and os.path.exists(png_file):
                print(f"✅ 转换成功: {png_file}")
                return True
            else:
                print(f"❌ 命令失败: {result.stderr}")
                
        except FileNotFoundError:
            print(f"❌ 命令不存在: {cmd[0]}")
        except Exception as e:
            print(f"❌ 转换错误: {e}")
    
    return False

if __name__ == "__main__":
    print("EMF to PNG 转换工具")
    print("=" * 30)
    
    if convert_emf():
        print("\\n🎉 转换完成！您现在可以查看PNG格式的流程图。")
    else:
        print("\\n💡 转换失败。建议:")
        print("1. 安装ImageMagick: brew install imagemagick")
        print("2. 或在Windows/Mac上手动打开EMF文件并另存为PNG")
'''
    
    script_path = "Files/test_embedded_objects/convert_to_png.py"
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write(converter_script)
    
    print(f"📝 已创建转换脚本: {script_path}")
    return script_path

def main():
    print("🎯 Visio流程图提取工具")
    print("=" * 40)
    
    # 创建HTML预览
    html_path = emf_to_base64()
    
    # 创建转换脚本
    script_path = create_simple_converter()
    
    print("\n📋 已为您准备的文件:")
    print(f"1. 📄 HTML预览文件: {html_path}")
    print(f"2. 🔧 转换脚本: {script_path}")
    print(f"3. 🖼️  EMF原文件: Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf")
    
    print("\n🚀 使用方法:")
    print("1. 用浏览器打开HTML文件查看流程图信息")
    print("2. 下载EMF文件到本地用其他工具转换")
    print("3. 或运行转换脚本尝试自动转换")

if __name__ == "__main__":
    main()
