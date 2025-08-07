import os
import zipfile
import shutil
import logging
import subprocess

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

def extract_embedded_files(docx_path, output_dir):
    """从 docx 中提取所有嵌入对象到 output_dir/embeddings"""
    os.makedirs(output_dir, exist_ok=True)
    embeddings_dir = os.path.join(output_dir, "embeddings")
    os.makedirs(embeddings_dir, exist_ok=True)
    with zipfile.ZipFile(docx_path, 'r') as docx:
        for file in docx.namelist():
            if file.startswith('word/embeddings/'):
                dst = os.path.join(embeddings_dir, os.path.basename(file))
                with docx.open(file) as src, open(dst, 'wb') as out:
                    out.write(src.read())
                logging.info(f"提取嵌入对象: {dst}")
    return embeddings_dir

def convert_visio_to_png(visio_path, png_path):
    """尝试用 libreoffice 或其他工具转换 vsdx/emf 为 png"""
    ext = os.path.splitext(visio_path)[-1].lower()
    # 只处理 vsdx/emf
    if ext == ".vsdx":
        try:
            result = subprocess.run([
                "soffice", "--headless", "--convert-to", "png", visio_path, "--outdir", os.path.dirname(png_path)
            ], capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                # 找到输出文件
                png_file = os.path.join(os.path.dirname(png_path), os.path.splitext(os.path.basename(visio_path))[0] + ".png")
                if os.path.exists(png_file):
                    shutil.move(png_file, png_path)
                    logging.info(f"VSDX转换成功: {png_path}")
                    return True
            logging.warning(f"LibreOffice转换失败: {result.stderr}")
        except Exception as e:
            logging.warning(f"LibreOffice转换异常: {e}")
    elif ext == ".emf":
        # macOS 可用 sips/qlmanage，Linux/Win 可提示手动或在线转换
        try:
            result = subprocess.run([
                "sips", "-s", "format", "png", visio_path, "--out", png_path
            ], capture_output=True, text=True, timeout=30)
            if result.returncode == 0 and os.path.exists(png_path):
                logging.info(f"EMF转换成功: {png_path}")
                return True
            else:
                logging.warning("sips转换失败，尝试qlmanage")
                result = subprocess.run([
                    "qlmanage", "-t", "-s", "1024", "-o", os.path.dirname(png_path), visio_path
                ], capture_output=True, text=True, timeout=30)
                # qlmanage 输出文件名可能不同
                for f in os.listdir(os.path.dirname(png_path)):
                    if f.endswith(".emf.png"):
                        shutil.move(os.path.join(os.path.dirname(png_path), f), png_path)
                        logging.info(f"qlmanage转换成功: {png_path}")
                        return True
        except Exception as e:
            logging.warning(f"EMF转换异常: {e}")
    return False

def batch_extract_docx_visio(docx_folder, output_base_dir):
    """批量处理文件夹下所有 docx，提取 Visio 并转为 png"""
    os.makedirs(output_base_dir, exist_ok=True)
    for fname in os.listdir(docx_folder):
        if not fname.lower().endswith('.docx') or fname.startswith('~$'):
            continue
        docx_path = os.path.join(docx_folder, fname)
        logging.info(f"处理文件: {docx_path}")
        file_output_dir = os.path.join(output_base_dir, os.path.splitext(fname)[0])
        embeddings_dir = extract_embedded_files(docx_path, file_output_dir)
        # 批量处理嵌入对象
        for embed_fname in os.listdir(embeddings_dir):
            ext = os.path.splitext(embed_fname)[-1].lower()
            if ext in [".vsd", ".vsdx", ".emf"]:
                visio_path = os.path.join(embeddings_dir, embed_fname)
                png_path = os.path.join(file_output_dir, embed_fname + ".png")
                success = convert_visio_to_png(visio_path, png_path)
                if not success:
                    logging.warning(f"未能自动转换: {visio_path}，可使用在线转换 https://convertio.co/emf-png/")
    logging.info("批量处理完成！")

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("用法: python extract_visio_to_png.py <docx_folder> <output_base_dir>")
        sys.exit(1)
    batch_extract_docx_visio(sys.argv[1], sys.argv[2])