#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Visio EMF 流程图处理器

专门用于处理从Visio嵌入对象中提取的EMF文件
提供多种转换和分析方法
"""

import os
import json
import subprocess
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class VisioEMFProcessor:
    """Visio EMF 文件处理器"""
    
    def __init__(self):
        self.supported_formats = ['png', 'jpg', 'svg', 'pdf']
        self.conversion_methods = {
            'inkscape': self._try_inkscape,
            'imagemagick': self._try_imagemagick,
            'sips': self._try_sips,
            'qlmanage': self._try_qlmanage
        }
    
    def process_visio_emf(self, emf_path, output_dir=None, target_format='png'):
        """
        处理Visio EMF文件
        
        Args:
            emf_path (str): EMF文件路径
            output_dir (str): 输出目录，默认为EMF文件所在目录
            target_format (str): 目标格式，默认PNG
            
        Returns:
            dict: 处理结果
        """
        emf_path = Path(emf_path)
        if not emf_path.exists():
            raise FileNotFoundError(f"EMF文件不存在: {emf_path}")
        
        if output_dir is None:
            output_dir = emf_path.parent
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        
        # 生成输出文件名
        base_name = emf_path.stem
        output_file = output_dir / f"{base_name}.{target_format}"
        
        # 尝试多种转换方法
        conversion_results = {}
        successful_conversion = None
        
        for method_name, method_func in self.conversion_methods.items():
            try:
                logger.info(f"尝试使用 {method_name} 转换 {emf_path.name}")
                result = method_func(emf_path, output_file)
                conversion_results[method_name] = result
                
                if result['success'] and successful_conversion is None:
                    successful_conversion = method_name
                    logger.info(f"✅ {method_name} 转换成功")
                    break
                    
            except Exception as e:
                conversion_results[method_name] = {
                    'success': False,
                    'error': str(e),
                    'output_file': None
                }
                logger.warning(f"❌ {method_name} 转换失败: {e}")
        
        # 分析EMF文件信息
        emf_info = self._analyze_emf_file(emf_path)
        
        # 生成处理报告
        report = {
            'emf_file': str(emf_path),
            'emf_info': emf_info,
            'target_format': target_format,
            'output_file': str(output_file) if successful_conversion else None,
            'successful_method': successful_conversion,
            'conversion_attempts': conversion_results,
            'timestamp': datetime.now().isoformat(),
            'status': 'success' if successful_conversion else 'failed',
            'recommendations': self._generate_recommendations(emf_info, conversion_results)
        }
        
        # 保存报告
        report_file = output_dir / f"{base_name}_conversion_report.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        logger.info(f"处理报告已保存: {report_file}")
        return report
    
    def _try_inkscape(self, emf_path, output_file):
        """尝试使用Inkscape转换"""
        try:
            cmd = [
                'inkscape',
                '--export-type=png' if output_file.suffix.lower() == '.png' else f'--export-type={output_file.suffix[1:]}',
                '--export-dpi=300',
                '--export-background=white',
                f'--export-filename={output_file}',
                str(emf_path)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and output_file.exists():
                return {
                    'success': True,
                    'output_file': str(output_file),
                    'method': 'inkscape',
                    'quality': 'high',
                    'dpi': 300
                }
            else:
                return {
                    'success': False,
                    'error': result.stderr or 'Unknown error',
                    'output_file': None
                }
                
        except subprocess.TimeoutExpired:
            return {'success': False, 'error': 'Inkscape转换超时', 'output_file': None}
        except FileNotFoundError:
            return {'success': False, 'error': 'Inkscape未安装', 'output_file': None}
    
    def _try_imagemagick(self, emf_path, output_file):
        """尝试使用ImageMagick转换"""
        try:
            cmd = [
                'convert',
                str(emf_path),
                '-density', '300',
                '-background', 'white',
                '-flatten',
                str(output_file)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and output_file.exists():
                return {
                    'success': True,
                    'output_file': str(output_file),
                    'method': 'imagemagick',
                    'quality': 'medium',
                    'dpi': 300
                }
            else:
                return {
                    'success': False,
                    'error': result.stderr or 'Unknown error',
                    'output_file': None
                }
                
        except subprocess.TimeoutExpired:
            return {'success': False, 'error': 'ImageMagick转换超时', 'output_file': None}
        except FileNotFoundError:
            return {'success': False, 'error': 'ImageMagick未安装', 'output_file': None}
    
    def _try_sips(self, emf_path, output_file):
        """尝试使用macOS sips转换"""
        try:
            cmd = [
                'sips',
                '-s', 'format', output_file.suffix[1:],
                '-s', 'dpiHeight', '300',
                '-s', 'dpiWidth', '300',
                str(emf_path),
                '--out', str(output_file)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and output_file.exists():
                return {
                    'success': True,
                    'output_file': str(output_file),
                    'method': 'sips',
                    'quality': 'medium',
                    'platform': 'macOS'
                }
            else:
                return {
                    'success': False,
                    'error': result.stderr or 'Unknown error',
                    'output_file': None
                }
                
        except subprocess.TimeoutExpired:
            return {'success': False, 'error': 'sips转换超时', 'output_file': None}
        except FileNotFoundError:
            return {'success': False, 'error': 'sips未找到（非macOS系统）', 'output_file': None}
    
    def _try_qlmanage(self, emf_path, output_file):
        """尝试使用macOS qlmanage生成预览"""
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                cmd = [
                    'qlmanage',
                    '-t',
                    '-s', '1024',
                    '-o', temp_dir,
                    str(emf_path)
                ]
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                
                # qlmanage生成的文件名格式为: filename.emf.png
                temp_output = Path(temp_dir) / f"{emf_path.name}.png"
                
                if result.returncode == 0 and temp_output.exists():
                    # 移动到目标位置
                    shutil.move(str(temp_output), str(output_file))
                    return {
                        'success': True,
                        'output_file': str(output_file),
                        'method': 'qlmanage',
                        'quality': 'medium',
                        'platform': 'macOS'
                    }
                else:
                    return {
                        'success': False,
                        'error': result.stderr or 'qlmanage转换失败',
                        'output_file': None
                    }
                    
        except subprocess.TimeoutExpired:
            return {'success': False, 'error': 'qlmanage转换超时', 'output_file': None}
        except FileNotFoundError:
            return {'success': False, 'error': 'qlmanage未找到（非macOS系统）', 'output_file': None}
    
    def _analyze_emf_file(self, emf_path):
        """分析EMF文件信息"""
        file_stat = emf_path.stat()
        
        info = {
            'file_size_bytes': file_stat.st_size,
            'file_size_kb': round(file_stat.st_size / 1024, 2),
            'creation_time': datetime.fromtimestamp(file_stat.st_ctime).isoformat(),
            'modification_time': datetime.fromtimestamp(file_stat.st_mtime).isoformat(),
            'file_type': 'Enhanced Metafile (EMF)',
            'likely_source': 'Microsoft Visio',
            'complexity_estimate': self._estimate_complexity(file_stat.st_size)
        }
        
        return info
    
    def _estimate_complexity(self, file_size):
        """根据文件大小估算复杂度"""
        if file_size < 10 * 1024:  # < 10KB
            return 'simple'
        elif file_size < 50 * 1024:  # < 50KB
            return 'medium'
        elif file_size < 200 * 1024:  # < 200KB
            return 'complex'
        else:
            return 'very_complex'
    
    def _generate_recommendations(self, emf_info, conversion_results):
        """生成处理建议"""
        recommendations = []
        
        # 基于文件大小的建议
        complexity = emf_info['complexity_estimate']
        if complexity in ['complex', 'very_complex']:
            recommendations.append("该Visio流程图较为复杂，建议使用专业软件（如Visio或LibreOffice Draw）打开查看")
        
        # 基于转换结果的建议
        successful_methods = [method for method, result in conversion_results.items() if result.get('success', False)]
        
        if not successful_methods:
            recommendations.extend([
                "自动转换失败，建议手动处理：",
                "1. 使用Microsoft Visio打开EMF文件并导出为PNG",
                "2. 使用在线转换工具：https://convertio.co/emf-png/",
                "3. 使用LibreOffice Draw导入EMF文件后导出",
                "4. 在Windows系统上双击EMF文件用默认程序打开"
            ])
        else:
            recommendations.append(f"推荐的转换方法：{', '.join(successful_methods)}")
        
        # 针对Visio特有的建议
        recommendations.extend([
            "Visio流程图通常包含丰富的业务逻辑，建议保留原始EMF文件作为参考",
            "如需要编辑流程图，建议在Visio中打开原始文档进行修改"
        ])
        
        return recommendations

def main():
    """主函数 - 用于测试和命令行使用"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Visio EMF 流程图处理器')
    parser.add_argument('emf_file', help='EMF文件路径')
    parser.add_argument('-o', '--output', help='输出目录')
    parser.add_argument('-f', '--format', default='png', choices=['png', 'jpg', 'svg', 'pdf'], help='输出格式')
    
    args = parser.parse_args()
    
    processor = VisioEMFProcessor()
    try:
        result = processor.process_visio_emf(
            emf_path=args.emf_file,
            output_dir=args.output,
            target_format=args.format
        )
        
        print(f"\n处理结果:")
        print(f"状态: {result['status']}")
        if result['successful_method']:
            print(f"成功方法: {result['successful_method']}")
            print(f"输出文件: {result['output_file']}")
        else:
            print("转换失败，请查看建议的手动处理方法")
        
        print(f"\n建议:")
        for rec in result['recommendations']:
            print(f"  • {rec}")
            
    except Exception as e:
        logger.error(f"处理失败: {e}")
        return 1
    
    return 0

if __name__ == '__main__':
    exit(main())
