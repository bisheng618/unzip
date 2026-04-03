import os
import sys
import traceback
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QRadioButton, QPushButton, QLineEdit, 
                             QLabel, QFileDialog, QMessageBox, QTextEdit, QGroupBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# 导入现有的业务逻辑
try:
    import compress_busi_images
    import compress_technical_images
    import compress_core_pro
except ImportError:
    print("错误: 找不到相关的核心逻辑脚本")

class Worker(QThread):
    """异步处理线程，避免界面卡顿"""
    finished = pyqtSignal(bool, str)
    log = pyqtSignal(str)

    def __init__(self, mode, directory, do_compress, do_privacy, ratio):
        super().__init__()
        self.mode = mode
        self.directory = directory
        self.do_compress = do_compress
        self.do_privacy = do_privacy
        self.ratio = ratio
        self._stop_requested = False

    def stop(self):
        self._stop_requested = True

    def check_stop(self):
        return self._stop_requested

    def run(self):
        class Logger:
            def __init__(self, signal):
                self.signal = signal
            def write(self, text):
                if text.strip():
                    self.signal.emit(text.strip())
            def flush(self):
                pass

        try:
            original_stdout = sys.stdout
            sys.stdout = Logger(self.log)
            
            try:
                if self.mode == "busi":
                    self.log.emit("开始处理商务文件...")
                    # 商务模式依然使用原有的扫描逻辑，但如果用户取消了某些选项，可能需要警告或调整
                    # 这里为了简化，如果用户选择了常规模式，我们就调用对应的原有逻辑
                    # 但如果用户在商务模式下取消了隐私清理，原有脚本不支持，所以我们改用通用逻辑扫描
                    self.run_generic_scan("商务")
                
                elif self.mode == "tech":
                    self.log.emit("开始处理技术文件...")
                    self.run_generic_scan("技术")
                    
                elif self.mode == "other":
                    self.log.emit("开始处理其他文件...")
                    self.run_generic_scan("其他")
            finally:
                sys.stdout = original_stdout
            
            if self._stop_requested:
                self.finished.emit(True, "处理已由用户终止。")
            else:
                self.finished.emit(True, "处理完成！")
        except Exception as e:
            error_msg = traceback.format_exc()
            self.finished.emit(False, f"处理失败:\n{error_msg}")

    def run_generic_scan(self, keyword):
        """通用的扫描与处理逻辑"""
        processed_count = 0
        for root, dirs, files in os.walk(self.directory):
            if self._stop_requested: break
            
            for file in files:
                if self._stop_requested: break
                
                # 筛选条件
                is_target = False
                if keyword == "商务":
                    is_target = file.endswith('.docx') and ('商务' in file) and ('最终版' in file) and '上传版' in root
                elif keyword == "技术":
                    is_target = file.endswith('.docx') and ('技术' in file) and ('最终版' in file) and '上传版' in root
                else: # 其他
                    is_target = file.endswith('.docx') and not file.startswith('~$') and '_backup' not in file

                if is_target and not file.startswith('~$') and '_backup' not in file:
                    file_path = os.path.join(root, file)
                    # 调用核心处理逻辑 (compress_core_pro)
                    # 默认压缩比例 0.7
                    success = compress_core_pro.process_single_file(
                        file_path, 
                        ratio=self.ratio, 
                        skip_compression=not self.do_compress,
                        skip_privacy=not self.do_privacy,
                        stop_check_func=self.check_stop
                    )
                    if success:
                        processed_count += 1
        
        self.log.emit(f"\n--- 扫描结束，共处理 {processed_count} 个文件 ---")

class ImageCompressorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("图片压缩工具 V1.0")
        self.resize(600, 550)

        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 1. 文件类型选择区
        type_group = QGroupBox("文件类型选择")
        type_layout = QHBoxLayout()
        self.radio_busi = QRadioButton("商务文件")
        self.radio_tech = QRadioButton("技术文件")
        self.radio_other = QRadioButton("其他文件")
        self.radio_busi.setChecked(True)
        
        # 绑定状态变化事件
        self.radio_busi.toggled.connect(self.update_default_options)
        self.radio_tech.toggled.connect(self.update_default_options)
        self.radio_other.toggled.connect(self.update_default_options)
        
        type_layout.addWidget(self.radio_busi)
        type_layout.addWidget(self.radio_tech)
        type_layout.addWidget(self.radio_other)
        type_group.setLayout(type_layout)
        main_layout.addWidget(type_group)

        # 2. 功能选项区
        opt_group = QGroupBox("功能处理选项")
        opt_layout = QHBoxLayout()
        from PyQt5.QtWidgets import QCheckBox, QComboBox
        self.cb_compress = QCheckBox("启用图片压缩")
        self.cb_privacy = QCheckBox("启用隐私清理")
        self.cb_compress.setChecked(True)
        self.cb_privacy.setChecked(True)
        
        # 增加比例选择
        opt_layout.addWidget(self.cb_compress)
        opt_layout.addWidget(QLabel("比例:"))
        self.combo_ratio = QComboBox()
        self.combo_ratio.addItems(["50%", "60%", "70%", "80%", "90%"])
        self.combo_ratio.setCurrentIndex(2) # 默认 70%
        opt_layout.addWidget(self.combo_ratio)
        opt_layout.addSpacing(20)
        opt_layout.addWidget(self.cb_privacy)
        
        opt_group.setLayout(opt_layout)
        main_layout.addWidget(opt_group)
        
        self.cb_compress.toggled.connect(lambda checked: self.combo_ratio.setEnabled(checked))

        # 2. 目录选择区
        dir_group = QGroupBox("选择处理目录")
        dir_layout = QHBoxLayout()
        self.dir_input = QLineEdit()
        self.dir_input.setPlaceholderText("请选择包含 docx 文件的文件夹...")
        btn_browse = QPushButton("浏览...")
        btn_browse.clicked.connect(self.browse_directory)
        dir_layout.addWidget(self.dir_input)
        dir_layout.addWidget(btn_browse)
        dir_group.setLayout(dir_layout)
        main_layout.addWidget(dir_group)

        # 3. 控制按钮区
        btn_layout = QHBoxLayout()
        self.btn_run = QPushButton("开始处理")
        self.btn_run.setFixedHeight(40)
        self.btn_run.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.btn_run.clicked.connect(self.start_processing)
        
        self.btn_stop = QPushButton("停止处理")
        self.btn_stop.setFixedHeight(40)
        self.btn_stop.setEnabled(False)
        self.btn_stop.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")
        self.btn_stop.clicked.connect(self.stop_processing)
        
        btn_layout.addWidget(self.btn_run)
        btn_layout.addWidget(self.btn_stop)
        main_layout.addLayout(btn_layout)

        # 4. 日志输出区
        log_group = QGroupBox("处理进度")
        log_layout = QVBoxLayout()
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        # 设置深色背景和浅色文字
        self.log_output.setStyleSheet("""
            background-color: #1e1e1e; 
            color: #d4d4d4; 
            font-family: 'Consolas', 'Courier New', monospace;
            font-size: 12px;
            padding: 5px;
        """)
        log_layout.addWidget(self.log_output)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if directory:
            self.dir_input.setText(directory)

    def append_log(self, text):
        self.log_output.append(text)
        # 自动滚动到底部
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def update_default_options(self):
        """根据选择的文件类型更新默认选项状态"""
        if self.radio_other.isChecked():
            # 其他文件默认不强制全选，或者根据需要设置
            pass
        else:
            # 商务/技术默认全选
            self.cb_compress.setChecked(True)
            self.cb_privacy.setChecked(True)

    def start_processing(self):
        directory = self.dir_input.text().strip()
        if not directory or not os.path.exists(directory):
            QMessageBox.warning(self, "提示", "请先选择有效的目录！")
            return

        if self.radio_busi.isChecked(): mode = "busi"
        elif self.radio_tech.isChecked(): mode = "tech"
        else: mode = "other"
        
        do_compress = self.cb_compress.isChecked()
        do_privacy = self.cb_privacy.isChecked()
        ratio = float(self.combo_ratio.currentText().replace("%", "")) / 100.0
        
        if not do_compress and not do_privacy:
            QMessageBox.information(self, "提示", "请至少选择一项处理功能（压缩或清理）！")
            return

        # 禁用或启用界面元素
        self.set_ui_processing_state(True)
        self.log_output.clear()
        self.append_log(f"--- 任务启动 ---")
        self.append_log(f"目录: {directory}")
        self.append_log(f"模式: {'商务文件' if mode == 'busi' else '技术文件' if mode == 'tech' else '其他文件'}")
        self.append_log(f"配置: 压缩={do_compress} (比例 {int(ratio*100)}%), 隐私清理={do_privacy}")

        # 启动线程
        self.worker = Worker(mode, directory, do_compress, do_privacy, ratio)
        self.worker.log.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def stop_processing(self):
        if self.worker and self.worker.isRunning():
            self.append_log("\n[等待] 正在发出停止指令，当前文件处理完后将终止...")
            self.worker.stop()
            self.btn_stop.setEnabled(False)
            self.btn_stop.setText("停止中...")

    def set_ui_processing_state(self, is_processing):
        self.btn_run.setEnabled(not is_processing)
        self.btn_stop.setEnabled(is_processing)
        self.btn_stop.setText("停止处理" if is_processing else "停止处理")
        self.radio_busi.setEnabled(not is_processing)
        self.radio_tech.setEnabled(not is_processing)
        self.radio_other.setEnabled(not is_processing)
        self.cb_compress.setEnabled(not is_processing)
        self.combo_ratio.setEnabled(not is_processing and self.cb_compress.isChecked())
        self.cb_privacy.setEnabled(not is_processing)
        self.dir_input.setEnabled(not is_processing)

    def on_finished(self, success, message):
        self.set_ui_processing_state(False)
        if success:
            self.append_log(f"\n[任务结束] {message}")
            QMessageBox.information(self, "提示", message)
        else:
            self.append_log(f"\n[错误] {message}")
            QMessageBox.critical(self, "错误", message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ImageCompressorGUI()
    window.show()
    sys.exit(app.exec_())
