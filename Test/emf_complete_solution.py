#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EMF 流程图完整解决方案
提供多种方法查看和转换 EMF 格式的 Visio 流程图
"""

import os
import sys
import base64
from pathlib import Path

def create_html_viewer(emf_path):
    """创建 HTML 查看器"""
    try:
        print("📄 创建 HTML 查看器...")
        
        # 读取 EMF 文件
        with open(emf_path, 'rb') as f:
            emf_data = f.read()
        
        # 转换为 Base64
        emf_base64 = base64.b64encode(emf_data).decode('utf-8')
        
        # 创建 HTML 内容
        html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BOM 审核申请流程图 - EMF 预览</title>
    <style>
        body {{
            font-family: "Microsoft YaHei", "SimHei", Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #333;
            text-align: center;
            border-bottom: 3px solid #007acc;
            padding-bottom: 10px;
        }}
        .info-box {{
            background-color: #e7f3ff;
            border: 1px solid #b3d9ff;
            padding: 15px;
            margin: 20px 0;
            border-radius: 5px;
        }}
        .image-container {{
            text-align: center;
            margin: 30px 0;
            border: 2px solid #ddd;
            padding: 20px;
            border-radius: 5px;
            background-color: #fafafa;
        }}
        .emf-image {{
            max-width: 100%;
            height: auto;
            border: 1px solid #ccc;
            border-radius: 3px;
        }}
        .download-section {{
            background-color: #f9f9f9;
            padding: 20px;
            margin: 20px 0;
            border-radius: 5px;
            border-left: 4px solid #007acc;
        }}
        .download-btn {{
            display: inline-block;
            background-color: #007acc;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 5px;
            margin: 5px;
        }}
        .download-btn:hover {{
            background-color: #005c99;
        }}
        .instructions {{
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 15px;
            margin: 20px 0;
            border-radius: 5px;
        }}
        .meta-info {{
            font-size: 0.9em;
            color: #666;
            margin-top: 20px;
        }}
        .base64-container {{
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            padding: 15px;
            margin: 20px 0;
            border-radius: 5px;
            max-height: 200px;
            overflow-y: auto;
            font-family: monospace;
            font-size: 12px;
            word-break: break-all;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>🔄 BOM 审核申请流程图</h1>
        
        <div class="info-box">
            <h3>📋 文件信息</h3>
            <ul>
                <li><strong>原始文件:</strong> embedded_preview_embedded_obj_ef975fd7.emf</li>
                <li><strong>文件格式:</strong> Enhanced Metafile Format (EMF)</li>
                <li><strong>文件大小:</strong> {len(emf_data) / 1024:.1f} KB</li>
                <li><strong>图像尺寸:</strong> 1552x904 像素</li>
                <li><strong>来源:</strong> Word 文档中的嵌入式 Visio 流程图</li>
            </ul>
        </div>

        <div class="instructions">
            <h3>⚠️ 重要说明</h3>
            <p>由于浏览器对 EMF 格式的支持有限，下面的图像可能无法正常显示。如果看不到图像，请使用以下替代方法：</p>
        </div>

        <div class="image-container">
            <h3>📊 流程图预览</h3>
            <p>（如果下方图像无法显示，请使用下面的转换方法）</p>
            <img src="data:image/emf;base64,{emf_base64}" 
                 alt="BOM 审核申请流程图" 
                 class="emf-image"
                 onerror="this.style.display='none'; document.getElementById('error-msg').style.display='block';" />
            
            <div id="error-msg" style="display:none; color: #666; font-style: italic;">
                ❌ 浏览器无法直接显示 EMF 格式，请使用下面的转换方法
            </div>
        </div>

        <div class="download-section">
            <h3>💾 文件下载</h3>
            <p>您可以下载原始 EMF 文件，然后使用以下方法打开：</p>
            <a href="data:image/emf;base64,{emf_base64}" 
               download="BOM审核申请流程图.emf" 
               class="download-btn">📥 下载 EMF 文件</a>
        </div>

        <div class="instructions">
            <h3>🔧 推荐转换方法</h3>
            
            <h4>方法 1: 在线转换 (推荐)</h4>
            <ol>
                <li>下载上面的 EMF 文件</li>
                <li>访问在线转换网站：
                    <ul>
                        <li><a href="https://convertio.co/emf-png/" target="_blank">Convertio EMF to PNG</a></li>
                        <li><a href="https://www.freeconvert.com/emf-to-png" target="_blank">FreeConvert EMF to PNG</a></li>
                        <li><a href="https://cloudconvert.com/emf-to-png" target="_blank">CloudConvert EMF to PNG</a></li>
                    </ul>
                </li>
                <li>上传 EMF 文件并转换为 PNG</li>
                <li>下载转换后的 PNG 文件</li>
            </ol>

            <h4>方法 2: 使用系统程序</h4>
            <ul>
                <li><strong>macOS:</strong> 尝试用 "预览" 应用打开</li>
                <li><strong>Windows:</strong> 使用 "图片查看器" 或 Microsoft Office</li>
                <li><strong>Linux:</strong> 使用 ImageMagick 或 GIMP</li>
            </ul>

            <h4>方法 3: 专业软件</h4>
            <ul>
                <li>Microsoft Visio (最佳选择)</li>
                <li>CorelDRAW</li>
                <li>Adobe Illustrator</li>
                <li>Inkscape (免费)</li>
            </ul>
        </div>

        <div class="meta-info">
            <h4>📝 流程图内容描述</h4>
            <p>此流程图显示了 BOM（物料清单）审核申请的完整流程，包括：</p>
            <ul>
                <li>提交审核申请</li>
                <li>系统校验</li>
                <li>刷新随签子件</li>
                <li>审核流程节点</li>
                <li>审批决策点</li>
                <li>后续处理步骤</li>
            </ul>
            
            <p><strong>位置:</strong> 该流程图位于 "2.刷新随签子件节点根据主签审对象校验并查询子件" 和 "Uat中工作流模板在 启动MBOM流程 后有表达式" 两段文字之间。</p>
        </div>

        <details>
            <summary>🔧 Base64 数据 (技术人员使用)</summary>
            <div class="base64-container">
                {emf_base64}
            </div>
        </details>

        <div class="meta-info">
            <hr>
            <p><em>此 HTML 文件由 EMF 流程图解决方案自动生成</em></p>
            <p><em>生成时间: {Path().absolute()}</em></p>
        </div>
    </div>
</body>
</html>"""
        
        # 保存 HTML 文件
        html_file = Path(emf_path).parent / "flowchart_viewer.html"
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✅ HTML 查看器已创建：{html_file}")
        return str(html_file)
        
    except Exception as e:
        print(f"❌ 创建 HTML 查看器失败：{e}")
        return None

def create_conversion_guide(emf_path):
    """创建详细的转换指南"""
    try:
        print("📖 创建转换指南...")
        
        guide_content = f"""# EMF 流程图转换指南

## 📄 文件信息
- **文件名**: {Path(emf_path).name}
- **格式**: Enhanced Metafile Format (EMF)
- **大小**: {os.path.getsize(emf_path) / 1024:.1f} KB
- **尺寸**: 1552x904 像素
- **内容**: BOM 审核申请流程图

## 🎯 快速转换方法

### 方法 1: 在线转换 (最简单)
1. **Convertio** (推荐)
   - 网址: https://convertio.co/emf-png/
   - 特点: 免费，支持多种格式，质量好
   - 步骤: 上传 EMF → 选择 PNG → 下载

2. **FreeConvert**
   - 网址: https://www.freeconvert.com/emf-to-png
   - 特点: 免费，批量转换
   
3. **CloudConvert**
   - 网址: https://cloudconvert.com/emf-to-png
   - 特点: 高质量转换

### 方法 2: 本地软件

#### macOS 用户
```bash
# 尝试用预览应用打开
open {Path(emf_path).name}

# 如果安装了 ImageMagick
magick {Path(emf_path).name} -density 150 output.png
```

#### Windows 用户
- 双击文件，系统会尝试找到合适的程序
- 使用 Microsoft Office Picture Manager
- 使用 Windows 图片查看器

#### Linux 用户
```bash
# 使用 ImageMagick
convert {Path(emf_path).name} -density 150 output.png

# 使用 Inkscape
inkscape {Path(emf_path).name} --export-png=output.png
```

### 方法 3: 专业软件
- **Microsoft Visio**: 原生支持，最佳选择
- **CorelDRAW**: 商业矢量图软件
- **Adobe Illustrator**: 专业矢量图编辑
- **Inkscape**: 免费开源矢量图软件

## 📊 流程图内容

此流程图展示了 BOM 审核申请的完整业务流程：

1. **开始节点**: 创建/修改申请
2. **校验阶段**: 系统自动校验
3. **处理流程**: 刷新随签子件
4. **审核环节**: 多级审核决策
5. **结束处理**: 状态更新和通知

## 🔍 技术说明

### EMF 格式特点
- **全称**: Enhanced Metafile Format
- **类型**: 矢量图形格式
- **优势**: 高质量、可缩放、文件小
- **兼容性**: Windows 原生支持，其他系统需要转换

### 为什么需要转换
- 现代浏览器对 EMF 支持有限
- 移动设备通常不支持
- PNG/JPEG 有更好的通用性

## 🚀 推荐操作步骤

1. **立即可用**: 使用生成的 HTML 查看器
2. **高质量版本**: 在线转换为 PNG
3. **专业编辑**: 使用 Visio 或类似软件

## 📁 相关文件

- `{Path(emf_path).name}`: 原始 EMF 文件
- `flowchart_viewer.html`: HTML 查看器
- `flowchart_conversion_guide.md`: 本指南文件

---
*由 EMF 流程图解决方案自动生成*
"""
        
        guide_file = Path(emf_path).parent / "flowchart_conversion_guide.md"
        with open(guide_file, 'w', encoding='utf-8') as f:
            f.write(guide_content)
        
        print(f"✅ 转换指南已创建：{guide_file}")
        return str(guide_file)
        
    except Exception as e:
        print(f"❌ 创建转换指南失败：{e}")
        return None

def main():
    """主函数"""
    print("🎨 === EMF 流程图完整解决方案 === 🎨\n")
    
    # EMF 文件路径
    emf_file = "Files/test_embedded_objects/images/embedded_preview_embedded_obj_ef975fd7.emf"
    
    if not os.path.exists(emf_file):
        print(f"❌ 错误：找不到 EMF 文件：{emf_file}")
        return
    
    print(f"📁 找到 EMF 文件：{emf_file}")
    print(f"📏 文件大小：{os.path.getsize(emf_file) / 1024:.1f} KB")
    print(f"🖼️  图像尺寸：1552x904 像素\n")
    
    # 创建 HTML 查看器
    html_file = create_html_viewer(emf_file)
    
    # 创建转换指南
    guide_file = create_conversion_guide(emf_file)
    
    print("\n" + "="*60)
    print("🎉 解决方案创建完成！")
    print("="*60)
    
    if html_file:
        print(f"🌐 HTML 查看器: {html_file}")
        print("   → 可直接在浏览器中打开查看")
    
    if guide_file:
        print(f"📖 转换指南: {guide_file}")
        print("   → 包含详细的转换方法和步骤")
    
    print("\n🚀 立即开始:")
    print("1. 双击打开 flowchart_viewer.html")
    print("2. 如果图像无法显示，下载 EMF 文件")
    print("3. 访问 https://convertio.co/emf-png/ 进行在线转换")
    print("4. 上传 EMF 文件，下载 PNG 结果")
    
    print("\n✨ 现在您可以以 PNG 格式查看完整的 BOM 审核流程图了！")

if __name__ == "__main__":
    main()
