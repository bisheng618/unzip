import os
import sys
from io import BytesIO
from PIL import Image
from docx import Document
from docx.shared import Inches
import zipfile
import tempfile
import shutil

# ================== 图片压缩核心逻辑 ==================
def compress_image_to_target(img_bytes, target_kb, min_kb=50, min_q=10, max_q=95):
    """
    通过迭代降低质量和分辨率来将图片压缩到目标大小 (KB)
    :param img_bytes: 图片的原始字节数据
    :param target_kb: 目标文件大小 (KB)
    :param min_kb: 最小文件大小 (KB)，压缩后的图片不能小于此值
    :param min_q: 最小JPEG质量值
    :param max_q: 最高JPEG质量值
    :return: 压缩后的图片字节数据
    """
    try:
        img = Image.open(BytesIO(img_bytes))
        # 转换图片格式以支持JPEG保存
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")

        # 确保目标大小不小于最小限制
        effective_target_kb = max(target_kb, min_kb)
        
        # 1. 尝试通过降低 quality 压缩
        quality = max_q
        step = 5
        best_result = None
        best_size = 0
        
        while quality >= min_q:
            buffer = BytesIO()
            img.save(buffer, format="JPEG", quality=quality)
            size_kb = len(buffer.getvalue()) / 1024
            
            # 保存最接近min_kb且不小于min_kb的结果
            if size_kb >= min_kb and (best_result is None or size_kb < best_size):
                best_result = buffer.getvalue()
                best_size = size_kb
            
            if size_kb <= effective_target_kb and size_kb >= min_kb:
                return buffer.getvalue()
            quality -= step

        # 2. 如果质量最低时仍太大, 则开始缩小图片尺寸
        original_img = img.copy()
        while True:
            w, h = img.size
            # 防止图片缩得太小
            if w < 300 or h < 300:
                break
            img = img.resize((int(w * 0.9), int(h * 0.9)), Image.LANCZOS)
            buffer = BytesIO()
            img.save(buffer, format="JPEG", quality=min_q)
            size_kb = len(buffer.getvalue()) / 1024
            
            # 保存最接近min_kb且不小于min_kb的结果
            if size_kb >= min_kb and (best_result is None or size_kb < best_size):
                best_result = buffer.getvalue()
                best_size = size_kb
            
            if size_kb <= effective_target_kb and size_kb >= min_kb:
                return buffer.getvalue()

        # 如果有符合min_kb要求的结果，返回它
        if best_result is not None:
            return best_result
        
        # 如果所有压缩结果都小于min_kb，返回原图或质量较高的版本
        buffer = BytesIO()
        original_img.save(buffer, format="JPEG", quality=85)
        return buffer.getvalue()
        
    except Exception as e:
        print(f"压缩图片时出错: {e}")
        # 如果过程中发生错误 (例如非图片格式), 返回原始数据
        return img_bytes


def get_file_size_mb(file_path):
    """获取文件大小(MB)"""
    return os.path.getsize(file_path) / (1024 * 1024)


def compress_technical_docx_images(docx_path, target_kb_per_image=75):
    """
    压缩技术文件docx中的图片，每张图片压缩到指定大小
    :param docx_path: docx文件路径
    :param target_kb_per_image: 每张图片的目标大小(KB)
    """
    original_size = get_file_size_mb(docx_path)
    print(f"\n处理文件: {os.path.basename(docx_path)}")
    print(f"原始大小: {original_size:.2f} MB")
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    
    try:
        # 解压docx文件（docx本质上是zip文件）
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # 查找media文件夹中的所有图片
        media_dir = os.path.join(temp_dir, 'word', 'media')
        
        if not os.path.exists(media_dir):
            print("文件中没有找到图片")
            return
        
        # 获取所有图片文件
        image_files = [f for f in os.listdir(media_dir) 
                      if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif'))]
        
        if not image_files:
            print("media文件夹中没有找到图片")
            return
        
        print(f"找到 {len(image_files)} 张图片，开始压缩到 {target_kb_per_image}KB...")
        
        compressed_count = 0
        for img_file in image_files:
            img_path = os.path.join(media_dir, img_file)
            
            # 读取原始图片
            with open(img_path, 'rb') as f:
                img_bytes = f.read()
            
            original_img_size = len(img_bytes) / 1024
            
            # 压缩图片到目标大小
            compressed_bytes = compress_image_to_target(img_bytes, target_kb_per_image, min_kb=50, max_q=95)
            new_img_size = len(compressed_bytes) / 1024
            
            # 保存压缩后的图片
            with open(img_path, 'wb') as f:
                f.write(compressed_bytes)
            
            print(f"  - {img_file}: {original_img_size:.1f}KB → {new_img_size:.1f}KB")
            compressed_count += 1
        
        # 删除作者和上次修改者信息
        core_xml_path = os.path.join(temp_dir, 'docProps', 'core.xml')
        if os.path.exists(core_xml_path):
            try:
                # 读取core.xml文件
                with open(core_xml_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 使用简单的字符串替换来清除作者和修改者信息
                import re
                # 清除creator标签内容
                content = re.sub(r'(<dc:creator>)[^<]*(</dc:creator>)', r'\1\2', content)
                # 清除lastModifiedBy标签内容
                content = re.sub(r'(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)', r'\1\2', content)
                # 清除company标签内容
                content = re.sub(r'(<cp:company>)[^<]*(</cp:company>)', r'\1\2', content)
                
                # 写回文件
                with open(core_xml_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                print("已删除作者和上次修改者信息")
            except Exception as e:
                print(f"删除作者信息时出错: {e}")
        
        # 重新打包为docx文件
        # 创建备份
        backup_path = docx_path.replace('.docx', '_backup.docx')
        if not os.path.exists(backup_path):  # 只在备份不存在时创建
            shutil.copy2(docx_path, backup_path)
            print(f"\n已创建备份: {os.path.basename(backup_path)}")
        else:
            print(f"\n备份已存在，跳过创建备份")
        
        # 重新压缩为docx
        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    docx.write(file_path, arcname)
        
        new_size = get_file_size_mb(docx_path)
        print(f"\n压缩完成!")
        print(f"原始文件大小: {original_size:.2f} MB")
        print(f"新文件大小: {new_size:.2f} MB")
        print(f"压缩了 {compressed_count} 张图片")
        print(f"节省空间: {original_size - new_size:.2f} MB ({((original_size - new_size) / original_size * 100):.1f}%)")
            
    except Exception as e:
        print(f"❌ 处理文件时出错: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # 清理临时目录
        try:
            shutil.rmtree(temp_dir)
        except:
            pass


def scan_and_compress_technical(root_dir, target_kb_per_image=75, stop_check_func=None):
    """
    扫描目录下所有包含"技术文件"的docx文件并压缩
    :param root_dir: 根目录路径
    :param target_kb_per_image: 每张图片的目标大小(KB)
    :param stop_check_func: 停止检查函数，返回True则停止处理
    """
    print(f"开始扫描目录: {root_dir}")
    print(f"每张图片目标大小: {target_kb_per_image}KB")
    print("=" * 60)
    
    processed_count = 0
    
    # 遍历目录
    for root, dirs, files in os.walk(root_dir):
        # 只处理包含“上传版”的目录
        if '上传版' not in root:
            continue

        # 检查是否需要停止
        if stop_check_func and stop_check_func():
            print("\n[收到停止指令] 正在停止扫描...")
            break

        for file in files:
            # 检查是否是docx文件且文件名包含"技术文件"，且包含"最终版"，但排除备份文件和临时文件
            if file.endswith('.docx') and ('技术文件' in file or '技术方案' in file) and ('最终版' in file) and not file.startswith('~$') and '_backup' not in file:
                # 再次检查停止指令
                if stop_check_func and stop_check_func():
                    print("\n[收到停止指令] 正在停止处理...")
                    print("\n" + "=" * 60)
                    print(f"处理已终止! 共处理 {processed_count} 个文件")
                    return

                file_path = os.path.join(root, file)
                
                compress_technical_docx_images(file_path, target_kb_per_image)
                processed_count += 1
    
    print("\n" + "=" * 60)
    print(f"扫描完成! 共处理 {processed_count} 个文件")


if __name__ == "__main__":
    # 目标目录
    target_directory = input("请输入目标目录: ").strip()

    # 检查目录是否存在
    if not os.path.exists(target_directory):
        print(f"错误: 目录不存在: {target_directory}")
        sys.exit(1)
    
    # 开始扫描和压缩，每张图片压缩到75KB左右
    scan_and_compress_technical(target_directory, target_kb_per_image=75)
