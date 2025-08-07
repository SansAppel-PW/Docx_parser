#!/usr/bin/env python3
"""
项目结构整理脚本
用于创建标准的项目目录结构和清理临时文件
"""

import os
import shutil
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def create_project_structure():
    """创建标准的项目目录结构"""
    
    directories = [
        "Files",
        "Files/examples",
        "parsed_docs",
        "parsed_docs/single_files", 
        "Test",
        "backups"
    ]
    
    logger.info("创建项目目录结构...")
    
    for directory in directories:
        try:
            os.makedirs(directory, exist_ok=True)
            logger.info(f"✅ 创建目录: {directory}")
        except Exception as e:
            logger.error(f"❌ 创建目录失败 {directory}: {e}")
    
    # 创建 .gitignore 文件
    gitignore_content = """# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg

# 日志文件
*.log

# 临时文件
temp*/
test_*_output/
quick_test_output/

# 解析结果（可选）
parsed_docs/
Files/structured_docs/
Files/single_file_output/

# VS Code
.vscode/

# macOS
.DS_Store

# Windows
Thumbs.db
"""
    
    try:
        with open('.gitignore', 'w', encoding='utf-8') as f:
            f.write(gitignore_content)
        logger.info("✅ 创建 .gitignore 文件")
    except Exception as e:
        logger.error(f"❌ 创建 .gitignore 失败: {e}")

def clean_temp_files():
    """清理临时文件和目录"""
    
    temp_patterns = [
        "__pycache__",
        "*.pyc",
        "temp*",
        "test_*_output",
        "quick_test_output",
        "test_performance_output"
    ]
    
    logger.info("清理临时文件...")
    
    # 删除缓存目录
    if os.path.exists("__pycache__"):
        try:
            shutil.rmtree("__pycache__")
            logger.info("✅ 删除 __pycache__ 目录")
        except Exception as e:
            logger.error(f"❌ 删除 __pycache__ 失败: {e}")
    
    # 清理临时输出目录
    temp_dirs = [
        "test_performance_output",
        "quick_test_output",
        "test_output_single"
    ]
    
    for temp_dir in temp_dirs:
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                logger.info(f"✅ 删除临时目录: {temp_dir}")
            except Exception as e:
                logger.error(f"❌ 删除临时目录失败 {temp_dir}: {e}")

def migrate_old_results():
    """迁移旧的解析结果到新的目录结构"""
    
    old_paths = [
        "Files/structured_docs",
        "Files/single_file_output"
    ]
    
    logger.info("迁移旧的解析结果...")
    
    for old_path in old_paths:
        if os.path.exists(old_path):
            # 创建备份
            backup_name = f"backup_{os.path.basename(old_path)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            backup_path = os.path.join("backups", backup_name)
            
            try:
                shutil.move(old_path, backup_path)
                logger.info(f"✅ 迁移 {old_path} 到 {backup_path}")
            except Exception as e:
                logger.error(f"❌ 迁移失败 {old_path}: {e}")

def organize_files():
    """整理文件到合适的目录"""
    
    # 如果当前目录有 DOCX 文件，建议移动到 Files 目录
    docx_files = [f for f in os.listdir('.') if f.lower().endswith('.docx') and not f.startswith('~$')]
    
    if docx_files:
        logger.info(f"发现 {len(docx_files)} 个DOCX文件，建议移动到Files目录")
        for docx_file in docx_files:
            target_path = os.path.join("Files", docx_file)
            if not os.path.exists(target_path):
                try:
                    shutil.move(docx_file, target_path)
                    logger.info(f"✅ 移动文件: {docx_file} -> Files/")
                except Exception as e:
                    logger.error(f"❌ 移动文件失败 {docx_file}: {e}")

def create_readme_if_missing():
    """如果没有README，创建一个基本的README"""
    
    if not os.path.exists("README.md"):
        readme_content = """# DOCX解析器

高性能的Word文档解析工具，支持提取文本、图片、SmartArt和嵌入对象。

## 快速开始

```bash
# 处理单个文件
python docx_parser.py Files/example.docx

# 批量处理文件夹
python docx_parser.py Files/PLM2.0

# 快速测试
python quick_parse_example.py
```

## 项目结构

- `Files/` - 存放原始Word文档
- `parsed_docs/` - 存放解析结果
- `Test/` - 测试脚本和开发工具

详细说明请查看 `项目结构说明.md`

## 特性

- 🚀 快速模式：跳过耗时的图像转换
- 📊 结构化输出：JSON格式的文档结构
- 🖼️ 图片提取：支持多种图片格式
- 🎨 SmartArt解析：提取图表文本内容
- 📦 嵌入对象：检测Visio、Excel等嵌入内容
- 📈 批量处理：支持文件夹批量处理

## 性能

- 单个文档处理时间：< 1秒
- 相比原版性能提升：97%+
- 支持大批量文档处理
"""
        
        try:
            with open("README.md", 'w', encoding='utf-8') as f:
                f.write(readme_content)
            logger.info("✅ 创建 README.md 文件")
        except Exception as e:
            logger.error(f"❌ 创建 README.md 失败: {e}")

def show_project_status():
    """显示项目当前状态"""
    
    logger.info("\n" + "="*100)
    logger.info("📊 项目状态总结")
    logger.info("="*50)
    
    # 统计文件数量
    files_count = 0
    if os.path.exists("Files"):
        for root, dirs, files in os.walk("Files"):
            files_count += len([f for f in files if f.lower().endswith('.docx')])
    
    parsed_count = 0
    if os.path.exists("parsed_docs"):
        for root, dirs, files in os.walk("parsed_docs"):
            parsed_count += len([f for f in files if f == 'document.json'])
    
    logger.info(f"📁 原始文档数量: {files_count}")
    logger.info(f"📄 已解析文档数量: {parsed_count}")
    
    # 检查关键文件
    key_files = [
        "docx_parser.py",
        "quick_parse_example.py", 
        "项目结构说明.md",
        "速度优化说明.md"
    ]
    
    logger.info("\n📋 关键文件检查:")
    for file in key_files:
        status = "✅" if os.path.exists(file) else "❌"
        logger.info(f"{status} {file}")
    
    # 检查目录结构
    required_dirs = ["Files", "parsed_docs", "Test"]
    logger.info("\n📁 目录结构检查:")
    for directory in required_dirs:
        status = "✅" if os.path.exists(directory) else "❌"
        logger.info(f"{status} {directory}/")
    
    logger.info("="*50)

def main():
    """主函数"""
    
    logger.info("🚀 开始整理项目结构...")
    
    # 1. 创建标准目录结构
    create_project_structure()
    
    # 2. 清理临时文件
    clean_temp_files()
    
    # 3. 迁移旧结果
    migrate_old_results()
    
    # 4. 整理文件
    organize_files()
    
    # 5. 创建README（如果缺失）
    create_readme_if_missing()
    
    # 6. 显示项目状态
    show_project_status()
    
    logger.info("✨ 项目结构整理完成！")
    logger.info("\n📝 接下来您可以:")
    logger.info("1. 将Word文档放入 Files/ 目录")
    logger.info("2. 运行 python docx_parser.py Files/your_folder 进行解析")
    logger.info("3. 查看 parsed_docs/ 目录中的解析结果")

if __name__ == "__main__":
    main()
