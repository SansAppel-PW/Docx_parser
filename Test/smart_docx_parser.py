#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能的自动化 DOCX 解析器，集成 EMF 到 PNG 的完整转换流程
"""

import os
import sys
import json
import subprocess
import logging
from pathlib import Path

# 添加当前目录到路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from docx_parser import parse_docx

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class SmartDocxParser:
    """智能 DOCX 解析器，包含自动 EMF 转换功能"""
    
    def __init__(self, enable_emf_conversion=True, conversion_timeout=30):
        self.enable_emf_conversion = enable_emf_conversion
        self.conversion_timeout = conversion_timeout
        self.conversion_stats = {
            'emf_found': 0,
            'emf_converted': 0,
            'conversion_methods': []
        }
    
    def parse_document(self, docx_path, output_dir):
        """解析 DOCX 文档"""
        logger.info(f"开始解析文档：{docx_path}")
        
        # 使用现有解析器
        result = parse_docx(docx_path, output_dir)
        
        if not result:
            logger.error("文档解析失败")
            return None
        
        # 如果启用了 EMF 转换，进行后处理
        if self.enable_emf_conversion:
            self._post_process_emf_files(output_dir, result)
        
        return result
    
    def _post_process_emf_files(self, output_dir, document_data):
        """后处理 EMF 文件，尝试转换为 PNG"""
        logger.info("开始 EMF 文件后处理...")
        
        images_dir = os.path.join(output_dir, "images")
        if not os.path.exists(images_dir):
            return
        
        # 查找所有 EMF 文件
        emf_files = [f for f in os.listdir(images_dir) if f.lower().endswith('.emf')]
        
        if not emf_files:
            logger.info("未发现 EMF 文件")
            return
        
        logger.info(f"发现 {len(emf_files)} 个 EMF 文件，开始转换...")
        self.conversion_stats['emf_found'] = len(emf_files)
        
        for emf_file in emf_files:
            emf_path = os.path.join(images_dir, emf_file)
            success = self._convert_emf_to_png(emf_path)
            
            if success:
                self.conversion_stats['emf_converted'] += 1
                # 更新 document_data 中的引用
                self._update_document_references(document_data, emf_file, emf_file.replace('.emf', '.png'))
        
        logger.info(f"EMF 转换完成：{self.conversion_stats['emf_converted']}/{self.conversion_stats['emf_found']} 成功")
    
    def _convert_emf_to_png(self, emf_path):
        """尝试多种方法将 EMF 转换为 PNG"""
        base_path = emf_path.replace('.emf', '')
        png_path = f"{base_path}.png"
        
        logger.info(f"尝试转换：{os.path.basename(emf_path)}")
        
        # 方法1: 使用 magick 命令（ImageMagick 7）
        if self._try_magick_convert(emf_path, png_path):
            return True
        
        # 方法2: 使用 convert 命令（ImageMagick 6）
        if self._try_convert_command(emf_path, png_path):
            return True
        
        # 方法3: 使用 Python Wand（如果可用）
        if self._try_wand_convert(emf_path, png_path):
            return True
        
        # 方法4: 使用 sips（macOS 系统工具）
        if self._try_sips_convert(emf_path, png_path):
            return True
        
        logger.warning(f"所有转换方法失败：{os.path.basename(emf_path)}")
        return False
    
    def _try_magick_convert(self, emf_path, png_path):
        """尝试使用 magick 命令转换"""
        try:
            # 设置环境变量
            env = os.environ.copy()
            env['MAGICK_HOME'] = '/opt/homebrew'
            env['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + env.get('DYLD_LIBRARY_PATH', '')
            
            cmd = ['magick', emf_path, '-density', '150', '-background', 'white', '-flatten', png_path]
            
            result = subprocess.run(cmd, 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=self.conversion_timeout,
                                  env=env)
            
            if result.returncode == 0 and os.path.exists(png_path) and os.path.getsize(png_path) > 0:
                logger.info(f"✅ magick 转换成功：{os.path.basename(png_path)}")
                self.conversion_stats['conversion_methods'].append('magick')
                return True
            else:
                logger.debug(f"magick 转换失败：{result.stderr}")
                
        except subprocess.TimeoutExpired:
            logger.warning(f"magick 转换超时：{os.path.basename(emf_path)}")
        except FileNotFoundError:
            logger.debug("magick 命令未找到")
        except Exception as e:
            logger.debug(f"magick 转换异常：{e}")
        
        return False
    
    def _try_convert_command(self, emf_path, png_path):
        """尝试使用 convert 命令转换"""
        try:
            env = os.environ.copy()
            env['MAGICK_HOME'] = '/opt/homebrew'
            env['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + env.get('DYLD_LIBRARY_PATH', '')
            
            cmd = ['convert', emf_path, '-density', '150', '-background', 'white', '-flatten', png_path]
            
            result = subprocess.run(cmd, 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=self.conversion_timeout,
                                  env=env)
            
            if result.returncode == 0 and os.path.exists(png_path) and os.path.getsize(png_path) > 0:
                logger.info(f"✅ convert 转换成功：{os.path.basename(png_path)}")
                self.conversion_stats['conversion_methods'].append('convert')
                return True
            else:
                logger.debug(f"convert 转换失败：{result.stderr}")
                
        except subprocess.TimeoutExpired:
            logger.warning(f"convert 转换超时：{os.path.basename(emf_path)}")
        except FileNotFoundError:
            logger.debug("convert 命令未找到")
        except Exception as e:
            logger.debug(f"convert 转换异常：{e}")
        
        return False
    
    def _try_wand_convert(self, emf_path, png_path):
        """尝试使用 Wand 转换"""
        try:
            # 设置环境变量
            os.environ['MAGICK_HOME'] = '/opt/homebrew'
            os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + os.environ.get('DYLD_LIBRARY_PATH', '')
            
            from wand.image import Image as WandImage
            
            with WandImage() as img:
                img.resolution = (150, 150)
                img.read(filename=emf_path)
                img.background_color = 'white'
                img.alpha_channel = 'remove'
                img.format = 'png'
                img.save(filename=png_path)
            
            if os.path.exists(png_path) and os.path.getsize(png_path) > 0:
                logger.info(f"✅ Wand 转换成功：{os.path.basename(png_path)}")
                self.conversion_stats['conversion_methods'].append('wand')
                return True
                
        except ImportError:
            logger.debug("Wand 库未安装")
        except Exception as e:
            logger.debug(f"Wand 转换失败：{e}")
        
        return False
    
    def _try_sips_convert(self, emf_path, png_path):
        """尝试使用 sips 转换（macOS）"""
        try:
            cmd = ['sips', '-s', 'format', 'png', emf_path, '--out', png_path]
            
            result = subprocess.run(cmd, 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=self.conversion_timeout)
            
            if result.returncode == 0 and os.path.exists(png_path) and os.path.getsize(png_path) > 0:
                logger.info(f"✅ sips 转换成功：{os.path.basename(png_path)}")
                self.conversion_stats['conversion_methods'].append('sips')
                return True
            else:
                logger.debug(f"sips 转换失败：{result.stderr}")
                
        except subprocess.TimeoutExpired:
            logger.warning(f"sips 转换超时：{os.path.basename(emf_path)}")
        except FileNotFoundError:
            logger.debug("sips 命令未找到")
        except Exception as e:
            logger.debug(f"sips 转换异常：{e}")
        
        return False
    
    def _update_document_references(self, document_data, old_filename, new_filename):
        """更新文档数据中的文件引用"""
        def update_refs(obj):
            if isinstance(obj, dict):
                if obj.get('type') == 'embedded_object' and obj.get('preview_image') == f"images/{old_filename}":
                    obj['preview_image'] = f"images/{new_filename}"
                    obj['converted_to_png'] = True
                    logger.debug(f"更新引用：{old_filename} -> {new_filename}")
                
                for value in obj.values():
                    update_refs(value)
            elif isinstance(obj, list):
                for item in obj:
                    update_refs(item)
        
        update_refs(document_data)
    
    def get_conversion_report(self):
        """获取转换报告"""
        return {
            'emf_files_found': self.conversion_stats['emf_found'],
            'emf_files_converted': self.conversion_stats['emf_converted'],
            'conversion_success_rate': f"{(self.conversion_stats['emf_converted'] / max(1, self.conversion_stats['emf_found']) * 100):.1f}%",
            'methods_used': list(set(self.conversion_stats['conversion_methods']))
        }

def main():
    """主函数"""
    print("🚀 === 智能 DOCX 解析器（自动 EMF 转换）=== 🚀\n")
    
    # 输入文件
    docx_file = "Files/PLM2.0/BOM 审核申请.docx"
    output_dir = "Files/smart_parser_output"
    
    if not os.path.exists(docx_file):
        print(f"❌ 文件不存在：{docx_file}")
        return
    
    # 创建智能解析器
    parser = SmartDocxParser(enable_emf_conversion=True)
    
    # 解析文档
    result = parser.parse_document(docx_file, output_dir)
    
    if result:
        # 保存结果
        result_file = os.path.join(output_dir, "document.json")
        with open(result_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        # 获取转换报告
        report = parser.get_conversion_report()
        
        print("✅ 解析完成！")
        print(f"📄 结果文件：{result_file}")
        print(f"\n📊 EMF 转换报告：")
        print(f"   - 发现 EMF 文件：{report['emf_files_found']} 个")
        print(f"   - 成功转换：{report['emf_files_converted']} 个")
        print(f"   - 成功率：{report['conversion_success_rate']}")
        
        if report['methods_used']:
            print(f"   - 使用方法：{', '.join(report['methods_used'])}")
        
        # 检查最终结果
        images_dir = os.path.join(output_dir, "images")
        if os.path.exists(images_dir):
            png_files = [f for f in os.listdir(images_dir) if f.endswith('.png') and 'embedded_preview' in f]
            emf_files = [f for f in os.listdir(images_dir) if f.endswith('.emf') and 'embedded_preview' in f]
            
            if png_files:
                print(f"\n🎉 转换成功的流程图：")
                for png_file in png_files:
                    print(f"   - {png_file}")
            
            if emf_files:
                print(f"\n⚠️  仍需手动处理的 EMF 文件：")
                for emf_file in emf_files:
                    print(f"   - {emf_file}")
                    print(f"     建议：使用在线转换工具 https://convertio.co/emf-png/")
        
        print(f"\n✨ 您现在可以直接在结果中查看 PNG 格式的流程图了！")
    
    else:
        print("❌ 解析失败")

if __name__ == "__main__":
    main()
