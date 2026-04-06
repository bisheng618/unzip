import sys
import os
import openpyxl
import shutil
from collections import defaultdict
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel,
                             QVBoxLayout, QHBoxLayout, QFileDialog, QTextEdit,
                             QMessageBox, QProgressBar)


class WorkerThread(QThread):
    """
    后台处理线程，负责文件扫描、匹配与复制。
    避免在大文件夹处理时造成界面卡死。
    """
    progress_signal = pyqtSignal(str)
    percent_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(int, dict, list)
    error_signal = pyqtSignal(str)

    def __init__(self, excel_file, source_dir, target_dir):
        super().__init__()
        self.excel_file = excel_file
        self.source_dir = source_dir
        self.target_dir = target_dir

    def run(self):
        try:
            # 1. 读取 Excel 文件 (改为使用 openpyxl)
            self.progress_signal.emit("正在读取 Excel 文件...")
            wb = openpyxl.load_workbook(self.excel_file, data_only=True)
            sheet = wb.active
            
            mobile_numbers = []
            header_col = -1
            
            # 在第一行寻找“号码”列
            for cell in sheet[1]:
                if cell.value == "号码":
                    header_col = cell.column
                    break
            
            if header_col == -1:
                self.error_signal.emit('Excel 文件首行中没有找到"号码"列')
                return
            
            # 读取该列所有手机号 (从第2行开始)
            for row in range(2, sheet.max_row + 1):
                val = sheet.cell(row=row, column=header_col).value
                if val is not None:
                    mobile_numbers.append(str(val).strip())

            if not mobile_numbers:
                self.error_signal.emit('Excel 文件中没有找到手机号数据')
                return

            # 2. 预扫描源文件夹建立索引
            self.progress_signal.emit("正在预扫描源文件夹建立索引...")
            all_files = os.listdir(self.source_dir)
            files_indexed = []
            for f in all_files:
                # 仅存储文件名，提高后续匹配效率
                files_indexed.append(f)

            copied_files = 0
            no_match_numbers = []
            mobile_file_count = defaultdict(int)

            # 3. 开始匹配和处理
            total = len(mobile_numbers)
            self.progress_signal.emit(f"开始匹配手机号 (总数: {total})...")

            for i, mobile_number in enumerate(mobile_numbers):
                found_match = False
                # 只需一次性遍历已建立的文件名列表，比频繁读盘快得多
                for filename in files_indexed:
                    if mobile_number in filename:
                        source_file = os.path.join(self.source_dir, filename)
                        target_file = os.path.join(self.target_dir, filename)

                        try:
                            shutil.copy2(source_file, target_file)
                            copied_files += 1
                            mobile_file_count[mobile_number] += 1
                            found_match = True
                        except Exception as e:
                            self.progress_signal.emit(f"复制文件 {filename} 时出错：{e}")

                if not found_match:
                    no_match_numbers.append(mobile_number)

                percent = int((i + 1) / total * 100)
                if percent % 5 == 0 or i + 1 == total:
                    self.progress_signal.emit(f"进度: {i + 1}/{total} ({percent}%)")
                    self.percent_signal.emit(percent)

            self.finished_signal.emit(copied_files, dict(mobile_file_count), no_match_numbers)

        except Exception as e:
            self.error_signal.emit(f"处理时发生非预期错误: {str(e)}")


class MobileFileSorterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('手机号文件批量处理 (性能优化版)')
        self.setGeometry(100, 100, 600, 500)

        # 布局
        layout = QVBoxLayout()

        # Excel文件选择
        excel_layout = QHBoxLayout()
        self.excel_label = QLabel('未选择Excel文件')
        excel_btn = QPushButton('选择Excel文件')
        excel_btn.clicked.connect(self.select_excel_file)
        excel_layout.addWidget(self.excel_label)
        excel_layout.addWidget(excel_btn)
        layout.addLayout(excel_layout)

        # 源文件夹选择
        source_layout = QHBoxLayout()
        self.source_label = QLabel('未选择源文件夹')
        source_btn = QPushButton('选择源文件夹')
        source_btn.clicked.connect(self.select_source_folder)
        source_layout.addWidget(self.source_label)
        source_layout.addWidget(source_btn)
        layout.addLayout(source_layout)

        # 目标文件夹选择
        target_layout = QHBoxLayout()
        self.target_label = QLabel('未选择目标文件夹')
        target_btn = QPushButton('选择目标文件夹')
        target_btn.clicked.connect(self.select_target_folder)
        target_layout.addWidget(self.target_label)
        target_layout.addWidget(target_btn)
        layout.addLayout(target_layout)

        # 开始处理按钮
        self.start_btn = QPushButton('开始处理')
        self.start_btn.clicked.connect(self.start_processing)
        layout.addWidget(self.start_btn)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat('%p%')
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 结果显示区域
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        layout.addWidget(self.result_text)

        self.setLayout(layout)

        # 存储文件路径
        self.excel_file = None
        self.source_directory = None
        self.target_directory = None

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择Excel文件', '', 'Excel Files (*.xlsx *.xls)')
        if file_path:
            self.excel_file = file_path
            self.excel_label.setText(f'已选择：{os.path.basename(file_path)}')

    def select_source_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, '选择源文件夹')
        if folder_path:
            self.source_directory = folder_path
            self.source_label.setText(f'已选择：{os.path.basename(folder_path)}')

    def select_target_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, '选择目标文件夹')
        if folder_path:
            self.target_directory = folder_path
            self.target_label.setText(f'已选择：{os.path.basename(folder_path)}')

    def start_processing(self):
        if not all([self.excel_file, self.source_directory, self.target_directory]):
            QMessageBox.warning(self, '错误', '请选择Excel文件、源文件夹和目标文件夹')
            return

        self.result_text.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.start_btn.setEnabled(False)  # 处理期间禁用按钮

        # 创建并启动后台线程
        self.worker = WorkerThread(self.excel_file, self.source_directory, self.target_directory)
        self.worker.progress_signal.connect(self.update_log)
        self.worker.percent_signal.connect(self.progress_bar.setValue)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error_signal.connect(self.on_error)
        self.worker.start()

    def update_log(self, message):
        self.result_text.append(message)

    def on_error(self, message):
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, '错误', message)
        self.start_btn.setEnabled(True)

    def on_finished(self, copied_files, mobile_file_count, no_match_numbers):
        self.result_text.append(f"\n--- 处理完成 ---")
        self.result_text.append(f"总共复制了 {copied_files} 个文件\n")

        # 打印多于1个文件的手机号
        multi_msg = [f"手机号 {num}: {count} 个文件" for num, count in mobile_file_count.items() if count > 1]
        if multi_msg:
            self.result_text.append("以下手机号有多个文件：")
            for msg in multi_msg:
                self.result_text.append(msg)

        # 打印没有匹配文件的手机号
        if no_match_numbers:
            self.result_text.append("\n以下手机号没有找到对应文件：")
            for number in no_match_numbers:
                self.result_text.append(number)

        self.progress_bar.setVisible(False)
        self.start_btn.setEnabled(True)
        QMessageBox.information(self, '完成', '文件处理已完成')


def main():
    app = QApplication(sys.argv)
    ex = MobileFileSorterApp()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()