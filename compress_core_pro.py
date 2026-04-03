import os
import sys
import zipfile
import tempfile
import shutil
import re
from io import BytesIO
from PIL import Image

def compress_image_by_ratio(img_bytes, ratio):
    """
    按比例压缩图片质量和尺寸
    :param img_bytes: 图片原始字节
    :param ratio: 压缩比例 (如 0.8, 0.7, 0.6, 0.5)
    :return: 压缩后的图片字节
    """
    try:
        img = Image.open(BytesIO(img_bytes))
        # 转换模式以支持 JPEG
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        
        # 1. 降低分辨率 (根据比例缩小宽高)
        # 注意：这里我们取比例的平方根作为边长的缩小因子，使得面积缩小到目标比例
        import math
        scale = math.sqrt(ratio)
        w, h = img.size
        new_w, new_h = int(w * scale), int(h * scale)
        if new_w > 10 and new_h > 10:
            img = img.resize((new_w, new_h), Image.LANCZOS)
        
        # 2. 降低质量 (将比例映射到质量分数，例如 80% 对应 quality=80)
        quality = int(ratio * 100)
        buffer = BytesIO()
        img.save(buffer, format="JPEG", quality=quality, optimize=True)
        return buffer.getvalue()
    except Exception as e:
        print(f"压缩图片时出错: {e}")
        return img_bytes

def clean_personal_info(temp_dir):
    """
    深度清除 docx 中的个人信息、修订记录和批注
    """
    # 1. 清理 core.xml (作者、修改者、日期)
    core_xml_path = os.path.join(temp_dir, 'docProps', 'core.xml')
    if os.path.exists(core_xml_path):
        try:
            with open(core_xml_path, 'r', encoding='utf-8') as f:
                content = f.read()
            # 清除 creator, lastModifiedBy, cp:company
            content = re.sub(r'(<dc:creator>)[^<]*(</dc:creator>)', r'\1\2', content)
            content = re.sub(r'(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)', r'\1\2', content)
            content = re.sub(r'(<cp:company>)[^<]*(</cp:company>)', r'\1\2', content)
            # 清除修改时间 (可选，但通常有助于隐私)
            # content = re.sub(r'(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)', r'\1\2', content)
            with open(core_xml_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print("  - 已清理 core.xml (作者信息)")
        except Exception as e:
            print(f"清理 core.xml 出错: {e}")

    # 2. 清理 app.xml (公司信息)
    app_xml_path = os.path.join(temp_dir, 'docProps', 'app.xml')
    if os.path.exists(app_xml_path):
        try:
            with open(app_xml_path, 'r', encoding='utf-8') as f:
                content = f.read()
            content = re.sub(r'(<Company>)[^<]*(</Company>)', r'\1\2', content)
            with open(app_xml_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print("  - 已清理 app.xml (公司信息)")
        except Exception as e:
            print(f"清理 app.xml 出错: {e}")

    # 3. 清理修订记录和批注 (Word 内部文件)
    # word/comments.xml, word/revisions.xml, word/settings.xml (trackRevisions)
    word_dir = os.path.join(temp_dir, 'word')
    if os.path.exists(word_dir):
        # 删除批注文件
        for extra in ['comments.xml', 'commentsExtended.xml', 'commentsIds.xml', 'people.xml']:
            extra_path = os.path.join(word_dir, extra)
            if os.path.exists(extra_path):
                os.remove(extra_path)
                print(f"  - 已删除 {extra} (批注/相关人员)")
        
        # 处理 settings.xml (关闭修订追踪开关)
        settings_path = os.path.join(word_dir, 'settings.xml')
        if os.path.exists(settings_path):
            try:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                # 移除 <w:trackRevisions/> 标签
                content = content.replace('<w:trackRevisions/>', '')
                with open(settings_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                print("  - 已在 settings.xml 中禁用修订追踪")
            except Exception as e:
                print(f"修改 settings.xml 出错: {e}")

        # 移除正文中的修订标记 (复杂但必要)
        # 这里使用正则表达式移除 <w:ins>, <w:del>, <w:commentRangeStart> 等
        # 注意：这只是一个基础实现，完全移除修订需要复杂的 XML 合并逻辑
        document_xml = os.path.join(word_dir, 'document.xml')
        if os.path.exists(document_xml):
            try:
                with open(document_xml, 'r', encoding='utf-8') as f:
                    content = f.read()
                # 移除修订相关的标签，但保留插入的内容 (w:ins) 里的文字
                # 这是一个简化的处理：将 <w:ins ...>内容</w:ins> 替换为 内容
                content = re.sub(r'<w:ins [^>]*>(.*?)</w:ins>', r'\1', content, flags=re.DOTALL)
                # 移除删除的内容 <w:del ...>内容</w:del>
                content = re.sub(r'<w:del [^>]*>.*?</w:del>', '', content, flags=re.DOTALL)
                # 移除批注引用
                content = re.sub(r'<w:commentRangeStart [^>]*/>', '', content)
                content = re.sub(r'<w:commentRangeEnd [^>]*/>', '', content)
                content = re.sub(r'<w:commentReference [^>]*/>', '', content)
                
                with open(document_xml, 'w', encoding='utf-8') as f:
                    f.write(content)
                print("  - 已清理 document.xml 中的修订标记和批注引用")
            except Exception as e:
                print(f"处理 document.xml 出错: {e}")

def process_single_file(file_path, ratio, skip_compression=False, skip_privacy=False, stop_check_func=None):
    """
    处理单个文件的压缩与清理
    :param skip_compression: 是否跳过图片压缩逻辑
    :param skip_privacy: 是否跳过隐私数据清理
    """
    print(f"\n[处理开始] 文件: {os.path.basename(file_path)}")
    
    modes = []
    if not skip_compression: modes.append(f"图片压缩 (比例: {int(ratio*100)}%)")
    if not skip_privacy: modes.append("隐私清理")
    
    if not modes:
        print("设定模式: 无需处理 (跳过所有步骤)")
        return True
    
    print(f"设定模式: {' + '.join(modes)}")
    
    if stop_check_func and stop_check_func(): return False

    temp_dir = tempfile.mkdtemp()
    try:
        # 解压
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # 1. 压缩图片 (仅在非跳过模式下执行)
        if not skip_compression:
            media_dir = os.path.join(temp_dir, 'word', 'media')
            if os.path.exists(media_dir):
                image_files = [f for f in os.listdir(media_dir) 
                              if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif'))]
                print(f"  - 找到 {len(image_files)} 张图片，开始处理...")
                for img_file in image_files:
                    if stop_check_func and stop_check_func(): break
                    img_path = os.path.join(media_dir, img_file)
                    with open(img_path, 'rb') as f:
                        img_bytes = f.read()
                    
                    compressed_bytes = compress_image_by_ratio(img_bytes, ratio)
                    with open(img_path, 'wb') as f:
                        f.write(compressed_bytes)
        else:
            print("  - 跳过图片压缩步骤")
        
        # 2. 清理个人信息
        if not skip_privacy:
            clean_personal_info(temp_dir)
        else:
            print("  - 跳过隐私清理步骤")

        # 3. 重新打包
        backup_path = file_path.replace('.docx', '_backup.docx')
        if not os.path.exists(backup_path):
            shutil.copy2(file_path, backup_path)
            print(f"  - 已备份原文件至: {os.path.basename(backup_path)}")
        
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    full_p = os.path.join(root, file)
                    arcname = os.path.relpath(full_p, temp_dir)
                    docx.write(full_p, arcname)
        
        print(f"[处理完成] {os.path.basename(file_path)}")
        return True
    except Exception as e:
        print(f"[处理失败] {e}")
        return False
    finally:
        shutil.rmtree(temp_dir)
