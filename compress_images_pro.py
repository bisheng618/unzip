import os
import sys
import zipfile
import tempfile
import shutil
import re
import math
import traceback
from io import BytesIO
from PIL import Image
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLineEdit, 
                             QLabel, QFileDialog, QMessageBox, QTextEdit, QGroupBox, QComboBox, QRadioButton)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# ================== 核心逻辑部分 (原 compress_core_pro.py) ==================

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
        
        # 1. 降低分辨率 (根据比例缩小面积)
        scale = math.sqrt(ratio)
        w, h = img.size
        new_w, new_h = int(w * scale), int(h * scale)
        if new_w > 10 and new_h > 10:
            img = img.resize((new_w, new_h), Image.LANCZOS)
        
        # 2. 降低质量
        quality = int(ratio * 100)
        buffer = BytesIO()
        img.save(buffer, format="JPEG", quality=quality, optimize=True)
        return buffer.getvalue()
    except Exception as e:
        print(f"压缩图片时出错: {e}")
        return img_bytes

def remove_unused_media(temp_dir):
    """
    清理 word/media 中未被引用的冗余图片
    """
    media_dir = os.path.join(temp_dir, 'word', 'media')
    if not os.path.exists(media_dir):
        return

    media_files = os.listdir(media_dir)
    if not media_files:
        return

    print("  - 检查冗余图片引用...")
    
    # 查找所有可能引用图片的 XML 和 rels 文件
    xml_files = []
    for root, dirs, files in os.walk(temp_dir):
        for f in files:
            if f.endswith(('.xml', '.rels')):
                xml_files.append(os.path.join(root, f))
    
    # 提前读取所有引用文件的内容
    all_content = ""
    for xml_p in xml_files:
        try:
            with open(xml_p, 'r', encoding='utf-8', errors='ignore') as f:
                all_content += f.read()
        except:
            continue

    removed_count = 0
    for m_file in media_files:
        # 在 XML 引用中查找该图片文件名 (basename)
        if m_file not in all_content:
            try:
                os.remove(os.path.join(media_dir, m_file))
                removed_count += 1
            except:
                pass
    
    if removed_count > 0:
        print(f"  - 已发现并清理 {removed_count} 个未引用的图片文件")
    else:
        print("  - 未发现冗余残留图片")

def clean_personal_info(temp_dir):
    """
    深度清除 docx 中的个人信息、修订记录和批注
    """
    # 1. 清理 core.xml (作者、修改者、公司)
    core_xml_path = os.path.join(temp_dir, 'docProps', 'core.xml')
    if os.path.exists(core_xml_path):
        try:
            with open(core_xml_path, 'r', encoding='utf-8') as f:
                content = f.read()
            content = re.sub(r'(<dc:creator>)[^<]*(</dc:creator>)', r'\1\2', content)
            content = re.sub(r'(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)', r'\1\2', content)
            content = re.sub(r'(<cp:company>)[^<]*(</cp:company>)', r'\1\2', content)
            with open(core_xml_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print("  - 已清理 core.xml (作者/上次修改者)")
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
            print("  - 已清理 app.xml (公司名)")
        except Exception as e:
            print(f"清理 app.xml 出错: {e}")

    # 3. 清理修订记录和批注
    word_dir = os.path.join(temp_dir, 'word')
    if os.path.exists(word_dir):
        # 删除批注依赖文件
        for extra in ['comments.xml', 'commentsExtended.xml', 'commentsIds.xml', 'people.xml', 'revisions.xml']:
            extra_path = os.path.join(word_dir, extra)
            if os.path.exists(extra_path):
                try:
                    os.remove(extra_path)
                    print(f"  - 已删除残留文件: {extra}")
                except:
                    pass
        
        # 处理 settings.xml (禁用修订追踪)
        settings_path = os.path.join(word_dir, 'settings.xml')
        if os.path.exists(settings_path):
            try:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                content = content.replace('<w:trackRevisions/>', '')
                with open(settings_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                print("  - 已在配置中禁用修订追踪开关")
            except Exception as e:
                print(f"修改 settings.xml 出错: {e}")

        # 移除正文中的修订标记和引用
        document_xml = os.path.join(word_dir, 'document.xml')
        if os.path.exists(document_xml):
            try:
                with open(document_xml, 'r', encoding='utf-8') as f:
                    content = f.read()
                # 简单移除修订和批注标签
                content = re.sub(r'<w:ins [^>]*>(.*?)</w:ins>', r'\1', content, flags=re.DOTALL)
                content = re.sub(r'<w:del [^>]*>.*?</w:del>', '', content, flags=re.DOTALL)
                content = re.sub(r'<w:commentRangeStart [^>]*/>', '', content)
                content = re.sub(r'<w:commentRangeEnd [^>]*/>', '', content)
                content = re.sub(r'<w:commentReference [^>]*/>', '', content)
                with open(document_xml, 'w', encoding='utf-8') as f:
                    f.write(content)
                print("  - 已清理正文中的修订痕迹与批注引用")
            except Exception as e:
                print(f"处理 document.xml 出错: {e}")

def process_single_file_logic(file_path, ratio, skip_compression=False, stop_check_func=None):
    """主处理逻辑"""
    print(f"\n[处理开始] 文件: {os.path.basename(file_path)}")
    if skip_compression:
        print("设定模式: 仅清除数据 (跳过图片压缩)")
    else:
        print(f"设定模式: 压缩+清除 (压缩率: {int(ratio*100)}%)")
    
    if stop_check_func and stop_check_func(): return False

    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # 1. 压缩图片
        if not skip_compression:
            media_dir = os.path.join(temp_dir, 'word', 'media')
            if os.path.exists(media_dir):
                image_files = [f for f in os.listdir(media_dir) 
                              if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif'))]
                print(f"  - 发现 {len(image_files)} 张图片，执行比例压缩...")
                for img_file in image_files:
                    if stop_check_func and stop_check_func(): break
                    img_path = os.path.join(media_dir, img_file)
                    with open(img_path, 'rb') as f:
                        img_bytes = f.read()
                    compressed_bytes = compress_image_by_ratio(img_bytes, ratio)
                    with open(img_path, 'wb') as f:
                        f.write(compressed_bytes)
        
        # 2. 清理冗余图片 (新增核心需求)
        remove_unused_media(temp_dir)

        # 3. 清理隐私数据
        clean_personal_info(temp_dir)

        # 4. 重新封包并备份
        backup_path = file_path.replace('.docx', '_backup.docx')
        if not os.path.exists(backup_path):
            shutil.copy2(file_path, backup_path)
            print(f"  - 已备份原始文件至 _backup.docx")
        
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    full_p = os.path.join(root, file)
                    arcname = os.path.relpath(full_p, temp_dir)
                    docx.write(full_p, arcname)
        
        print(f"[处理完成] 成功处理: {os.path.basename(file_path)}")
        return True
    except Exception as e:
        print(f"[处理失败] 错误详情: {e}")
        return False
    finally:
        shutil.rmtree(temp_dir)

# ================== GUI 部分 (原 compress_images_pro.py) ==================

class ProWorker(QThread):
    """异步处理线程"""
    finished = pyqtSignal(bool, str)
    log = pyqtSignal(str)

    def __init__(self, file_path, ratio, skip_compression=False):
        super().__init__()
        self.file_path = file_path
        self.ratio = ratio
        self.skip_compression = skip_compression
        self._stop_requested = False

    def stop(self):
        self._stop_requested = True

    def check_stop(self):
        return self._stop_requested

    def run(self):
        class Logger:
            def __init__(self, signal): self.signal = signal
            def write(self, text): 
                if text.strip(): self.signal.emit(text.strip())
            def flush(self): pass

        original_stdout = sys.stdout
        sys.stdout = Logger(self.log)
        try:
            success = process_single_file_logic(
                self.file_path, self.ratio, skip_compression=self.skip_compression, stop_check_func=self.check_stop
            )
            if self._stop_requested:
                self.finished.emit(True, "处理已手动终止。")
            elif success:
                self.finished.emit(True, "文件处理成功！")
            else:
                self.finished.emit(False, "处理过程中出现错误，请检查日志。")
        except Exception:
            self.finished.emit(False, f"程序运行时发生异常:\n{traceback.format_exc()}")
        finally:
            sys.stdout = original_stdout

class ImageCompressorProGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("图片压缩清理工具 Pro (单文件版)")
        self.resize(650, 680)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 1. 文件选择
        file_group = QGroupBox("选择 Word 文档")
        file_layout = QHBoxLayout()
        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText("选择具体的 .docx 文件进行处理...")
        btn_browse = QPushButton("选择文件...")
        btn_browse.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_input)
        file_layout.addWidget(btn_browse)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # 2. 处理模式
        mode_group = QGroupBox("处理模式选择")
        mode_layout = QHBoxLayout()
        self.radio_both = QRadioButton("压缩图片 + 去除冗余与隐私")
        self.radio_only_clean = QRadioButton("只去除冗余与隐私")
        self.radio_both.setChecked(True)
        self.radio_both.toggled.connect(self.on_mode_changed)
        mode_layout.addWidget(self.radio_both)
        mode_layout.addWidget(self.radio_only_clean)
        mode_group.setLayout(mode_layout)
        main_layout.addWidget(mode_group)

        # 3. 压缩设置
        param_group = QGroupBox("参数配置")
        param_layout = QHBoxLayout()
        param_layout.addWidget(QLabel("图片比例:"))
        self.combo_ratio = QComboBox()
        self.combo_ratio.addItems(["80%", "70%", "60%", "50%"])
        self.combo_ratio.setCurrentIndex(0)
        param_layout.addWidget(self.combo_ratio)
        param_layout.addStretch()
        
        info_label = QLabel("(包含: 作者、修订记录、批注、未引用图片)")
        info_label.setStyleSheet("color: #666; font-size: 11px;")
        param_layout.addWidget(info_label)
        param_group.setLayout(param_layout)
        main_layout.addWidget(param_group)

        # 4. 操作按钮
        btn_layout = QHBoxLayout()
        self.btn_run = QPushButton("立即开始处理")
        self.btn_run.setFixedHeight(45)
        self.btn_run.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; border-radius: 4px;")
        self.btn_run.clicked.connect(self.start_processing)
        
        self.btn_stop = QPushButton("终止计划")
        self.btn_stop.setFixedHeight(45)
        self.btn_stop.setEnabled(False)
        self.btn_stop.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; border-radius: 4px;")
        self.btn_stop.clicked.connect(self.stop_processing)
        
        btn_layout.addWidget(self.btn_run)
        btn_layout.addWidget(self.btn_stop)
        main_layout.addLayout(btn_layout)

        # 5. 回显日志
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout()
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("""
            background-color: #1e1e1e; 
            color: #d4d4d4; 
            font-family: 'Consolas', 'Menlo', monospace;
            font-size: 12px;
            padding: 8px;
        """)
        log_layout.addWidget(self.log_output)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

    def on_mode_changed(self):
        is_compress = self.radio_both.isChecked()
        self.combo_ratio.setEnabled(is_compress)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 Word 文件", "", "Word 文档 (*.docx)")
        if file_path:
            self.file_input.setText(file_path)

    def append_log(self, text):
        self.log_output.append(text)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def start_processing(self):
        file_path = self.file_input.text().strip()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "错误", "请先选择需要处理的 .docx 文件！")
            return

        ratio_str = self.combo_ratio.currentText().replace("%", "")
        ratio = float(ratio_str) / 100.0
        skip_compression = self.radio_only_clean.isChecked()
        
        self.set_ui_state(True)
        self.log_output.clear()
        self.append_log(f"--- 任务启动 ---")
        self.append_log(f"处理路径: {file_path}")
        
        self.worker = ProWorker(file_path, ratio, skip_compression=skip_compression)
        self.worker.log.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def stop_processing(self):
        if self.worker and self.worker.isRunning():
            self.append_log("\n[终止信号] 正在等待当前步骤结束以安全退出...")
            self.worker.stop()
            self.btn_stop.setEnabled(False)

    def set_ui_state(self, is_processing):
        self.btn_run.setEnabled(not is_processing)
        self.btn_stop.setEnabled(is_processing)
        self.file_input.setEnabled(not is_processing)
        self.radio_both.setEnabled(not is_processing)
        self.radio_only_clean.setEnabled(not is_processing)
        if not is_processing:
            self.on_mode_changed()
        else:
            self.combo_ratio.setEnabled(False)

    def on_finished(self, success, message):
        self.set_ui_state(False)
        if success:
            QMessageBox.information(self, "处理结束", message)
        else:
            QMessageBox.critical(self, "出错啦", message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ImageCompressorProGUI()
    window.show()
    sys.exit(app.exec_())
