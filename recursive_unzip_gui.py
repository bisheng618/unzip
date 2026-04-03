#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import zipfile
import shutil
import tarfile
import openpyxl
from openpyxl.styles import Alignment, Font
import copy
import subprocess
import re
import threading
import platform
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit, 
    QProgressBar, QMessageBox, QSplitter, QRadioButton, QButtonGroup, QGroupBox
)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QTextCursor

class UnzipWorker(QThread):
    """后台解压线程，避免界面卡顿"""
    progress_updated = pyqtSignal(str)
    status_changed = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, source_dir, target_dir, mode='full'):
        super().__init__()
        self.source_dir = source_dir
        self.target_dir = target_dir
        self.mode = mode
        self.running = True
    
    def run(self):
        try:
            # 根据模式执行不同的操作
            if self.mode == 'copy_engineering':
                # 仅复制工程量清单
                # 只需要检查目标目录
                if not os.path.exists(self.target_dir):
                    self.error.emit(f"目标目录不存在: {self.target_dir}")
                    return
                
                self.status_changed.emit("开始工程量清单组织...")
                self.copy_engineering_files(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始处理工程量清单报价表头...")
                self.process_engineering_summary_excel(self.target_dir)
                self.status_changed.emit("完成！")
                self.progress_updated.emit("工程量清单组织已完成！")
                
            elif self.mode == 'convert_doc':
                # 仅转换DOC为DOCX
                # 只需要检查目标目录
                if not os.path.exists(self.target_dir):
                    self.error.emit(f"目标目录不存在: {self.target_dir}")
                    return
                
                self.status_changed.emit("开始技术规范书.doc转.docx...")
                self.convert_doc_to_docx(self.target_dir)
                self.status_changed.emit("完成！")
                self.progress_updated.emit("DOC转DOCX已完成！")
                
            elif self.mode == 'check_technical':
                # 仅检查技术规范书
                # 只需要检查目标目录
                if not os.path.exists(self.target_dir):
                    self.error.emit(f"目标目录不存在: {self.target_dir}")
                    return
                
                self.status_changed.emit("检查缺少技术规范书的包...")
                self.check_package_technical_specs(self.target_dir)
                self.status_changed.emit("完成！")
                self.progress_updated.emit("技术规范书检查已完成！")
                
            elif self.mode == 'check_engineering':
                # 仅检查工作量清单
                # 只需要检查目标目录
                if not os.path.exists(self.target_dir):
                    self.error.emit(f"目标目录不存在: {self.target_dir}")
                    return
                
                self.status_changed.emit("检查缺少工作量清单的包...")
                self.check_package_engineering_specs(self.target_dir)
                self.status_changed.emit("完成！")
                self.progress_updated.emit("工作量清单检查已完成！")
                
            else:  # mode == 'full' 或其他值，执行完整流程
                # 检查源目录和目标目录
                if not os.path.exists(self.source_dir):
                    self.error.emit(f"源目录不存在: {self.source_dir}")
                    return
                
                if not os.path.exists(self.target_dir):
                    os.makedirs(self.target_dir)
                    self.progress_updated.emit(f"创建目标目录: {self.target_dir}")
                
                # 执行解压和处理操作
                self.status_changed.emit("开始初始解压...")
                self.initial_extraction()
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始递归扫描解压...")
                self.process_directory(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始文件组织...")
                self.organize_files(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始技术规范书组织...")
                self.organize_technical_specs(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始清理...")
                self.cleanup(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始处理SGCC文件...")
                self.process_sgcc_files(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始重命名包文件夹...")
                self.create_package_subfolders(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始工程量清单组织...")
                self.copy_engineering_files(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("开始处理工程量清单报价表头...")
                self.process_engineering_summary_excel(self.target_dir)

                if not self.running:
                    return
                
                self.status_changed.emit("开始删除.sign文件...")
                self.delete_sign_files(self.target_dir)
                
                if not self.running:
                    return
                
                # self.status_changed.emit("开始技术规范书.doc转.docx...")
                # self.convert_doc_to_docx(self.target_dir)
                
                # if not self.running:
                #     return
                
                self.status_changed.emit("检查缺少技术规范书的包...")
                self.check_package_technical_specs(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("检查缺少工作量清单的包...")
                self.check_package_engineering_specs(self.target_dir)
                
                if not self.running:
                    return
                
                self.status_changed.emit("完成！")
                self.progress_updated.emit("所有操作已完成！")
            
        except Exception as e:
            self.error.emit(f"执行过程中出错: {str(e)}")
        finally:
            self.finished.emit()
    
    def stop(self):
        self.running = False
    
    def log(self, message):
        """输出日志信息"""
        if self.running:
            self.progress_updated.emit(message)
    
    def unzip_file(self, zip_path, extract_to):
        """解压ZIP文件，处理中文文件名"""
        if not self.running:
            return
            
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    # 尝试修复编码：CP437 -> GBK
                    try:
                        filename = file_info.filename.encode('cp437').decode('gbk')
                    except:
                        filename = file_info.filename

                    # 构建完整输出路径
                    target_path = os.path.join(extract_to, filename)
                    
                    # 处理目录条目
                    if file_info.is_dir():
                        if not os.path.exists(target_path):
                            os.makedirs(target_path)
                        continue

                    # 确保父目录存在
                    parent_dir = os.path.dirname(target_path)
                    if not os.path.exists(parent_dir):
                        os.makedirs(parent_dir)

                    # 解压文件
                    with zip_ref.open(file_info) as source, open(target_path, "wb") as target:
                        shutil.copyfileobj(source, target)
                        
            self.log(f"解压: {zip_path} -> {extract_to}")
        except zipfile.BadZipFile:
            self.log(f"错误: 损坏的ZIP文件 {zip_path}")
        except Exception as e:
            self.log(f"解压 {zip_path} 时出错: {e}")
    
    def unrar_file(self, rar_path, extract_to):
        """使用unar解压RAR文件"""
        if not self.running:
            return
            
        try:
            # unar -o <output_dir> -D <archive>
            # -D: 不要为解压内容创建包含目录
            # -f: 强制覆盖
            cmd = ['unar', '-o', extract_to, '-D', '-f', rar_path]
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if result.returncode == 0:
                self.log(f"解压RAR: {rar_path} -> {extract_to}")
            else:
                self.log(f"解压RAR {rar_path} 时出错: {result.stderr}")
        except Exception as e:
            self.log(f"解压RAR {rar_path} 时出错: {e}")
    
    def un7z_file(self, sevenz_path, extract_to):
        """使用7z命令行工具解压.7z文件"""
        if not self.running:
            return
            
        try:
            # 7z x <archive> -o<output_dir> -y
            # x: 解压并保持目录结构
            # -o: 输出目录
            # -y: 对所有提示回答yes
            cmd = ['7z', 'x', sevenz_path, f'-o{extract_to}', '-y']
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if result.returncode == 0:
                self.log(f"解压7z: {sevenz_path} -> {extract_to}")
            else:
                self.log(f"解压7z {sevenz_path} 时出错: {result.stderr}")
        except Exception as e:
            self.log(f"解压7z {sevenz_path} 时出错: {e}")
    
    def untar_gz_file(self, tar_gz_path, extract_to):
        """使用tarfile模块解压.tar.gz文件"""
        if not self.running:
            return
            
        try:
            with tarfile.open(tar_gz_path, 'r:gz') as tar:
                tar.extractall(path=extract_to)
            self.log(f"解压tar.gz: {tar_gz_path} -> {extract_to}")
        except Exception as e:
            self.log(f"解压tar.gz {tar_gz_path} 时出错: {e}")
    
    def process_directory(self, directory):
        """递归扫描目录中的压缩文件并解压"""
        if not self.running:
            return
            
        for root, dirs, files in os.walk(directory):
            if not self.running:
                return
                
            for file in files:
                if not self.running:
                    return
                    
                file_lower = file.lower()
                if file_lower.endswith('.zip') or file_lower.endswith('.rar') or file_lower.endswith('.7z') or file_lower.endswith('.tar.gz'):
                    file_path = os.path.join(root, file)
                    # 创建与压缩文件同名的目录（无扩展名）
                    if file_lower.endswith('.tar.gz'):
                        # 对于.tar.gz，移除.tar.gz
                        extract_dir_name = file[:-7]
                    else:
                        extract_dir_name = os.path.splitext(file)[0]
                    extract_to = os.path.join(root, extract_dir_name)
                    
                    if not os.path.exists(extract_to):
                        os.makedirs(extract_to)
                        if file_lower.endswith('.zip'):
                            self.unzip_file(file_path, extract_to)
                        elif file_lower.endswith('.rar'):
                            self.unrar_file(file_path, extract_to)
                        elif file_lower.endswith('.7z'):
                            self.un7z_file(file_path, extract_to)
                        elif file_lower.endswith('.tar.gz'):
                            self.untar_gz_file(file_path, extract_to)
                        
                        # 递归处理新解压的目录
                        self.process_directory(extract_to)
    
    def initial_extraction(self):
        """初始解压：从源目录到目标目录"""
        for item in os.listdir(self.source_dir):
            if not self.running:
                return
                
            item_path = os.path.join(self.source_dir, item)
            item_lower = item.lower()
            if os.path.isfile(item_path) and (item_lower.endswith('.zip') or item_lower.endswith('.rar') or item_lower.endswith('.7z') or item_lower.endswith('.tar.gz')):
                if item_lower.endswith('.tar.gz'):
                    extract_dir_name = item[:-7]
                else:
                    extract_dir_name = os.path.splitext(item)[0]
                extract_to = os.path.join(self.target_dir, extract_dir_name)
                
                if not os.path.exists(extract_to):
                    os.makedirs(extract_to)
                
                if item_lower.endswith('.zip'):
                    self.unzip_file(item_path, extract_to)
                elif item_lower.endswith('.rar'):
                    self.unrar_file(item_path, extract_to)
                elif item_lower.endswith('.7z'):
                    self.un7z_file(item_path, extract_to)
                elif item_lower.endswith('.tar.gz'):
                    self.untar_gz_file(item_path, extract_to)
    
    def move_item(self, src, dst):
        """移动项目，处理冲突"""
        if not self.running:
            return
            
        if src == dst:
            return
            
        if os.path.exists(dst):
            self.log(f"目标已存在，跳过: {dst}")
            return
            
        try:
            shutil.move(src, dst)
            self.log(f"移动: {src} -> {dst}")
        except Exception as e:
            self.log(f"移动 {src} 到 {dst} 时出错: {e}")
    
    def organize_files(self, root_dir):
        """扫描包含"采购文件包"的目录并组织其中的文件"""
        if not self.running:
            return
            
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            if "采购文件包" in os.path.basename(dirpath):
                package_root = dirpath
                self.log(f"处理包: {package_root}")
                
                for sub_root, sub_dirs, sub_files in os.walk(package_root):
                    if not self.running:
                        return
                        
                    if sub_root == package_root:
                        continue  # 不移动已经在根目录的文件
                    
                    # 检查文件
                    for file in sub_files:
                        if not self.running:
                            return
                            
                        if ("应答注意" in file or "采购文件" in file):
                            src_path = os.path.join(sub_root, file)
                            dst_path = os.path.join(package_root, file)
                            self.move_item(src_path, dst_path)
                    
                    # 检查文件夹
                    for d in sub_dirs[:]:
                        if not self.running:
                            return
                            
                        if "采购公告" in d:
                            src_path = os.path.join(sub_root, d)
                            dst_path = os.path.join(package_root, d)
                            self.move_item(src_path, dst_path)
    
    def organize_technical_specs(self, root_dir):
        """扫描包含"技术规范书"的目录并组织文件"""
        if not self.running:
            return
            
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            if "技术规范书" in os.path.basename(dirpath):
                spec_dir = dirpath
                parent_dir = os.path.dirname(spec_dir)
                self.log(f"处理技术规范: {spec_dir} -> {parent_dir}")
                
                # 第一轮：将技术规范书目录中的所有文件移动到父目录
                for sub_root, sub_dirs, sub_files in os.walk(spec_dir):
                    if not self.running:
                        return
                        
                    for file in sub_files:
                        if not self.running:
                            return
                            
                        src_path = os.path.join(sub_root, file)
                        dst_path = os.path.join(parent_dir, file)
                        self.move_item(src_path, dst_path)
                
                # 第二轮：扫描父目录下所有目录，查找包含特定关键字的文件
                for sub_root, sub_dirs, sub_files in os.walk(parent_dir):
                    if not self.running:
                        return
                        
                    # 跳过技术规范书目录本身（已处理）
                    if sub_root.startswith(spec_dir):
                        continue
                        
                    for file in sub_files:
                        if not self.running:
                            return
                            
                        if "技术规范书" in file or "工程量清单" in file or "工作量清单" in file:
                            src_path = os.path.join(sub_root, file)
                            dst_path = os.path.join(parent_dir, file)
                            self.move_item(src_path, dst_path)
    
    def cleanup(self, root_dir):
        """清理操作：删除压缩文件和技术规范书目录"""
        if not self.running:
            return
            
        # 先删除文件
        for dirpath, dirnames, filenames in os.walk(root_dir, topdown=True):
            if not self.running:
                return
                
            # 删除zip、rar、7z和tar.gz文件
            for file in filenames:
                if not self.running:
                    return
                    
                file_lower = file.lower()
                if file_lower.endswith('.zip') or file_lower.endswith('.rar') or file_lower.endswith('.7z') or file_lower.endswith('.tar.gz'):
                    file_path = os.path.join(dirpath, file)
                    try:
                        os.remove(file_path)
                        self.log(f"删除压缩包: {file_path}")
                    except Exception as e:
                        self.log(f"删除 {file_path} 时出错: {e}")
        
        # 再删除目录
        for dirpath, dirnames, filenames in os.walk(root_dir, topdown=False):
            if not self.running:
                return
                
            # 删除"技术规范书"目录
            for d in dirnames[:]:
                if not self.running:
                    return
                    
                if "技术规范书" in d:
                    dir_path = os.path.join(dirpath, d)
                    try:
                        shutil.rmtree(dir_path)
                        self.log(f"删除技术规范目录: {dir_path}")
                    except Exception as e:
                        self.log(f"删除 {dir_path} 时出错: {e}")
    
    def process_sgcc_files(self, root_dir):
        """处理SGCC文件"""
        if not self.running:
            return
            
        sgcc_data = []
        
        # 1. 查找root_dir中的"采购文件包"目录
        package_roots = []
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            if "采购文件包" in os.path.basename(dirpath):
                package_roots.append(dirpath)
        
        # 如果没有找到采购文件包，也考虑root_dir本身
        search_roots = package_roots if package_roots else [root_dir]
        
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            if "SGCC包" in os.path.basename(dirpath):
                self.log(f"扫描SGCC包: {dirpath}")
                for file in filenames:
                    if not self.running:
                        return
                        
                    if file.endswith('.SGCC'):
                        parts = file.split('_')
                        row = [file] + parts
                        sgcc_data.append(row)

                        # 复制逻辑
                        if len(parts) >= 3:
                            part2 = parts[1]  # "Part 2"
                            part3 = parts[2]  # "Part 3"
                            
                            target_sub_1 = None
                            found_root = None
                            
                            # 尝试在search_roots中找到Part 2
                            for s_root in search_roots:
                                if not self.running:
                                    return
                                    
                                target_sub_1 = self.find_subdir_containing(s_root, part2)
                                if target_sub_1:
                                    found_root = s_root
                                    break
                            
                            if target_sub_1:
                                # 从target_sub_1中查找Part 3的子目录
                                target_sub_2 = None
                                search_terms = [part3]
                                
                                m = re.match(r'^包(\d+)$', part3)
                                if m:
                                    num = m.group(1)
                                    if len(num) == 1:
                                        search_terms.append(f"包0{num}")
                                
                                for term in search_terms:
                                    if not self.running:
                                        return
                                        
                                    target_sub_2 = self.find_subdir_containing(target_sub_1, term)
                                    if target_sub_2:
                                        break
                                
                                if target_sub_2:
                                    src_file = os.path.join(dirpath, file)
                                    dst_file = os.path.join(target_sub_2, file)
                                    try:
                                        shutil.copy2(src_file, dst_file)
                                        self.log(f"复制.SGCC文件: {file} -> {target_sub_2}")
                                    except Exception as e:
                                        self.log(f"复制.SGCC文件 {file} 时出错: {e}")
                                else:
                                    self.log(f"跳过复制: 无法在'{target_sub_1}'中找到匹配'{part3}'（或填充版本）的子目录")
                            else:
                                self.log(f"跳过复制: 无法在检查的根目录中找到匹配'{part2}'的目录")
        
        if sgcc_data:
            excel_path = os.path.join(root_dir, "SGCC_文件汇总.xlsx")
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "SGCC Files"
                
                # 表头
                header = ["原始文件名", "第1部分", "第2部分", "第3部分", "第4部分", "第5部分", "第6部分", "第7部分", "第8部分"]
                ws.append(header)
                
                for row in sgcc_data:
                    ws.append(row)
                    
                wb.save(excel_path)
                self.log(f"保存SGCC文件名到: {excel_path}")
            except Exception as e:
                self.log(f"保存Excel文件时出错: {e}")
        else:
            self.log("在'SGCC包'目录中未找到.SGCC文件")
    
    def find_subdir_containing(self, base_path, search_term):
        """在base_path中查找包含search_term的子目录"""
        if not os.path.exists(base_path):
            return None
        for item in os.listdir(base_path):
            if not self.running:
                return None
                
            full_path = os.path.join(base_path, item)
            if os.path.isdir(full_path) and search_term in item:
                return full_path
        return None
    
    def create_package_subfolders(self, root_dir):
        """在以"包"开头的文件夹中创建"上传版"和"生成版"子文件夹"""
        if not self.running:
            return
            
        import re
        
        # 在所有包文件夹中创建子文件夹
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            for dirname in dirnames:
                # 检查文件夹名称是否以"包"后跟数字开头
                if re.match(r'^包\d+', dirname):
                    package_path = os.path.join(dirpath, dirname)
                    upload_folder = os.path.join(package_path, "上传版")
                    generate_folder = os.path.join(package_path, "生成版")
                    
                    # 创建"上传版"文件夹
                    if not os.path.exists(upload_folder):
                        try:
                            os.makedirs(upload_folder)
                            self.log(f"创建文件夹: {upload_folder}")
                        except Exception as e:
                            self.log(f"创建 {upload_folder} 时出错: {e}")
                    
                    # 创建"生成版"文件夹
                    if not os.path.exists(generate_folder):
                        try:
                            os.makedirs(generate_folder)
                            self.log(f"创建文件夹: {generate_folder}")
                        except Exception as e:
                            self.log(f"创建 {generate_folder} 时出错: {e}")
    
    def copy_engineering_files(self, root_dir):
        """创建"工程量清单"文件夹并复制相关文件"""
        if not self.running:
            return
            
        target_folder = os.path.join(root_dir, "工程量清单_汇总")
        excel_summary_path = os.path.join(root_dir, "工作量清单_文件汇总.xlsx")
        
        # 1. 先删除"工作量清单_文件汇总.xlsx"文件
        if os.path.exists(excel_summary_path):
            try:
                os.remove(excel_summary_path)
                self.log(f"删除旧的汇总文件: {excel_summary_path}")
            except Exception as e:
                self.log(f"删除汇总文件时出错: {e}")
        
        # 2. 清空"工程量清单_汇总"文件夹
        if os.path.exists(target_folder):
            try:
                # 删除文件夹中的所有内容
                for item in os.listdir(target_folder):
                    item_path = os.path.join(target_folder, item)
                    if os.path.isfile(item_path):
                        os.remove(item_path)
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                self.log(f"已清空文件夹: {target_folder}")
            except Exception as e:
                self.log(f"清空文件夹时出错: {e}")
        else:
            # 如果文件夹不存在,则创建
            os.makedirs(target_folder)
            self.log(f"创建文件夹: {target_folder}")
            
        self.log(f"复制'工程量清单'和'工作量清单'文件到: {target_folder}")
        
        copied_files_info = []

        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            # 避免扫描目标文件夹本身
            if dirpath == target_folder:
                continue
                
            for file in filenames:
                if not self.running:
                    return
                    
                if "工程量清单" in file or "工作量清单" in file:
                    src_path = os.path.join(dirpath, file)
                    dst_path = os.path.join(target_folder, file)
                    
                    # 处理重复文件名
                    if os.path.exists(dst_path):
                        base, ext = os.path.splitext(file)
                        counter = 1
                        while os.path.exists(dst_path):
                            dst_path = os.path.join(target_folder, f"{base}_{counter}{ext}")
                            counter += 1
                    
                    try:
                        shutil.copy2(src_path, dst_path)
                        self.log(f"复制: {src_path} -> {dst_path}")
                        copied_files_info.append([os.path.basename(dst_path), dirpath])
                    except Exception as e:
                        self.log(f"复制 {src_path} 时出错: {e}")

        # 创建汇总Excel文件
        if copied_files_info:
            excel_path = os.path.join(root_dir, "工作量清单_文件汇总.xlsx")
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Engineering Files Summary"
                
                # 表头
                ws.append(["文件名", "文件路径"])
                
                for row in copied_files_info:
                    ws.append(row)
                    
                wb.save(excel_path)
                self.log(f"保存工程量文件汇总到: {excel_path}")
            except Exception as e:
                self.log(f"保存汇总Excel文件时出错: {e}")
    
    def delete_sign_files(self, root_dir):
        """递归删除所有以.sign结尾的文件"""
        if not self.running:
            return
            
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            for file in filenames:
                if not self.running:
                    return
                    
                if file.endswith('.sign'):
                    file_path = os.path.join(dirpath, file)
                    try:
                        os.remove(file_path)
                        self.log(f"删除.sign文件: {file_path}")
                    except Exception as e:
                        self.log(f"删除 {file_path} 时出错: {e}")
    
    def convert_doc_to_docx(self, root_dir):
        """将包含"技术规范书"或"技规书"的.doc文件转换为.docx格式"""
        if not self.running:
            return
            
        import subprocess
        
        converted_count = 0
        
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            for file in filenames:
                if not self.running:
                    return
                    
                # 检查文件是否为.doc且包含"技术规范书"、"技规书"或"技术文件规范"
                if file.endswith('.doc') and not file.endswith('.docx') and ("技术规范书" in file or "技规书" in file or "技术文件规范" in file):
                    doc_path = os.path.join(dirpath, file)
                    original_file = file  # 保存原始文件名用于日志
                    
                    # 调试日志：记录找到的文件
                    self.log(f"[调试] 找到待处理文件: {file}")
                    self.log(f"[调试] 包含'技术文件规范': {'技术文件规范' in file}, 包含'技术规范书': {'技术规范书' in file}, 包含'技规书': {'技规书' in file}")
                    
                    # 处理顺序很重要：先处理"技术文件规范"，再处理"技规书"
                    
                    # 第一步：如果文件名包含"技术文件规范"但不包含"技术规范书"，则在文件名后追加"技术规范书"
                    if "技术文件规范" in file and "技术规范书" not in file:
                        # 在扩展名前插入"技术规范书"
                        base_name = os.path.splitext(file)[0]
                        new_filename = base_name + "技术规范书.doc"
                        new_doc_path = os.path.join(dirpath, new_filename)
                        self.log(f"[调试] 准备重命名: {file} -> {new_filename}")
                        try:
                            os.rename(doc_path, new_doc_path)
                            self.log(f"✓ 重命名成功（添加技术规范书）: {file} -> {new_filename}")
                            doc_path = new_doc_path
                            file = new_filename
                        except Exception as e:
                            self.log(f"✗ 重命名失败: {file}")
                            self.log(f"  错误详情: {type(e).__name__}: {str(e)}")
                            continue
                    
                    # 第二步：如果文件名包含"技规书"但不包含"技术规范书"，替换为"技术规范书"
                    # 注意：经过第一步后，如果原来有"技术文件规范"，现在已经有"技术规范书"了，所以不会再执行这一步
                    if "技规书" in file and "技术规范书" not in file:
                        new_filename = file.replace("技规书", "技术规范书")
                        new_doc_path = os.path.join(dirpath, new_filename)
                        self.log(f"[调试] 准备重命名: {file} -> {new_filename}")
                        try:
                            os.rename(doc_path, new_doc_path)
                            self.log(f"✓ 重命名成功（技规书->技术规范书）: {file} -> {new_filename}")
                            doc_path = new_doc_path
                            file = new_filename
                        except Exception as e:
                            self.log(f"✗ 重命名失败: {file}")
                            self.log(f"  错误详情: {type(e).__name__}: {str(e)}")
                            continue
                    
                    docx_filename = os.path.splitext(file)[0] + '.docx'
                    docx_path = os.path.join(dirpath, docx_filename)
                    
                    # 如果.docx已存在则跳过
                    if os.path.exists(docx_path):
                        self.log(f"跳过转换（DOCX已存在）: {doc_path}")
                        continue
                    
                    try:
                        # 使用LibreOffice转换.doc为.docx
                        # --headless: 无GUI运行
                        # --convert-to docx: 转换为DOCX格式
                        # --outdir: 输出目录（与源文件相同）
                        cmd = [
                            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                            '--headless',
                            '--convert-to', 'docx',
                            '--outdir', dirpath,
                            doc_path
                        ]
                        
                        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=60)
                        
                        if result.returncode == 0 and os.path.exists(docx_path):
                            self.log(f"转换: {doc_path} -> {docx_path}")
                            converted_count += 1
                        else:
                            self.log(f"转换 {doc_path} 时出错: {result.stderr}")
                            
                    except subprocess.TimeoutExpired:
                        self.log(f"转换 {doc_path} 超时")
                    except Exception as e:
                        self.log(f"转换 {doc_path} 时出错: {e}")
        
        if converted_count > 0:
            self.log(f"总共转换了 {converted_count} 个.doc文件到.docx")
        else:
            self.log("没有找到包含'技术规范书'或'技规书'的.doc文件需要转换")
    
    def check_package_technical_specs(self, root_dir):
        """检查缺少技术规范书的包"""
        if not self.running:
            return
            
        missing_specs_dirs = []
        
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            # 检查当前目录名是否以"包"开头
            dir_basename = os.path.basename(dirpath)
            if dir_basename.startswith("包"):
                # 检查此目录中是否有包含"技术规范书"的文件
                has_tech_spec = False
                for file in filenames:
                    if "技术规范书" in file:
                        has_tech_spec = True
                        break
                
                # 如果未找到技术规范文件，记录此目录
                if not has_tech_spec:
                    missing_specs_dirs.append(dirpath)
                    self.log(f"缺少技术规范书: {dirpath}")
        
        # 写入结果到文件
        if missing_specs_dirs:
            output_file = os.path.join(root_dir, "缺少技术规范书的目录.txt")
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write("以下目录缺少技术规范书文件:\n")
                    f.write("=" * 60 + "\n\n")
                    for dir_path in missing_specs_dirs:
                        f.write(f"{dir_path}\n")
                    f.write("\n" + "=" * 60 + "\n")
                    f.write(f"总计: {len(missing_specs_dirs)} 个目录缺少技术规范书\n")
                self.log(f"\n已保存缺少技术规范书的报告到: {output_file}")
                self.log(f"缺少技术规范书的目录总数: {len(missing_specs_dirs)}")
            except Exception as e:
                self.log(f"保存报告文件时出错: {e}")
        else:
            self.log("所有以'包'开头的目录都包含技术规范书文件")
    
    def check_package_engineering_specs(self, root_dir):
        """检查缺少工作量清单的包"""
        if not self.running:
            return
            
        missing_engineering_dirs = []
        
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if not self.running:
                return
                
            # 检查当前目录名是否以"包"开头
            dir_basename = os.path.basename(dirpath)
            if dir_basename.startswith("包"):
                # 检查此目录中是否有包含"工作量清单"或"工程量清单"的文件
                has_engineering_spec = False
                for file in filenames:
                    if "工作量清单" in file or "工程量清单" in file:
                        has_engineering_spec = True
                        break
                
                # 如果未找到工程量清单文件，记录此目录
                if not has_engineering_spec:
                    missing_engineering_dirs.append([dirpath])
                    self.log(f"缺少工作量清单: {dirpath}")
        
        # 写入结果到Excel文件
        if missing_engineering_dirs:
            excel_path = os.path.join(root_dir, "缺少工作量清单的目录.xlsx")
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Missing Engineering Specs"
                
                # 表头
                ws.append(["目录路径"])
                
                for row in missing_engineering_dirs:
                    ws.append(row)
                    
                wb.save(excel_path)
                self.log(f"\n已保存缺少工作量清单的报告到: {excel_path}")
                self.log(f"缺少工作量清单的目录总数: {len(missing_engineering_dirs)}")
            except Exception as e:
                self.log(f"保存Excel报告文件时出错: {e}")
        else:
            self.log("所有以'包'开头的目录都包含工作量清单文件")

    def process_engineering_summary_excel(self, root_dir):
        """处理工程量清单汇总文件夹中的Excel文件"""
        if not self.running:
            return

        summary_dir = os.path.join(root_dir, "工程量清单_汇总")
        if not os.path.exists(summary_dir):
            self.log("未找到'工程量清单_汇总'文件夹，跳过Excel处理")
            return

        for filename in os.listdir(summary_dir):
            if not self.running:
                return
            
            # 处理 .xlsx 和 .xls 文件
            if (filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls')) and not filename.startswith('~$'):
                file_path = os.path.join(summary_dir, filename)
                
                # 如果是 .xls 文件，先使用 LibreOffice 转换为 .xlsx
                if filename.lower().endswith('.xls') and not filename.lower().endswith('.xlsx'):
                    try:
                        # 使用 LibreOffice 转换 .xls 为 .xlsx，保留所有格式
                        # --headless: 无GUI运行
                        # --convert-to xlsx: 转换为XLSX格式
                        # --outdir: 输出目录（与源文件相同）
                        cmd = [
                            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                            '--headless',
                            '--convert-to', 'xlsx',
                            '--outdir', summary_dir,
                            file_path
                        ]
                        
                        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=60)
                        
                        new_file_path = os.path.splitext(file_path)[0] + '.xlsx'
                        
                        if result.returncode == 0 and os.path.exists(new_file_path):
                            # 删除原 .xls 文件
                            os.remove(file_path)
                            
                            # 更新文件路径和文件名
                            file_path = new_file_path
                            filename = os.path.basename(new_file_path)
                            
                            self.log(f"已将 .xls 文件转换为 .xlsx: {filename}")
                        else:
                            self.log(f"转换 .xls 文件 {filename} 失败: {result.stderr}")
                            continue
                            
                    except subprocess.TimeoutExpired:
                        self.log(f"转换 .xls 文件 {filename} 超时")
                        continue
                    except Exception as e:
                        self.log(f"转换 .xls 文件 {filename} 时出错: {e}")
                        continue
                
                # 处理 .xlsx 文件（包括刚转换的）
                try:
                    wb = openpyxl.load_workbook(file_path)
                    # 总是处理第一个sheet
                    ws = wb.worksheets[0]
                    # 修改Sheet名为“可编辑版本”
                    ws.title = "可编辑版本"
                    
                    # 1. 修改第一行内容为“报价明细表”
                    # 假设这里是修改A1单元格，或者第一行的合并单元格
                    # 用户需求：将第一行的文字内容改为“报价明细表”，其他字体格式保持不变
                    # 通常表头在第一行
                    first_cell = ws.cell(row=1, column=1)
                    first_cell.value = "报价明细表"

                    
                    # 2. 在最后一行下面再增加一行，根据有几列进行合并单元格
                    # 寻找包含数据的实际最后一行（忽略仅有格式的空行）
                    max_row = ws.max_row
                    found_data = False
                    for r in range(max_row, 0, -1):
                        for c in range(1, ws.max_column + 1):
                            cell_val = ws.cell(row=r, column=c).value
                            if cell_val is not None and str(cell_val).strip() != "":
                                max_row = r
                                found_data = True
                                break
                        if found_data:
                            break
                    
                    if not found_data:
                        max_row = 1  # 只有表头的情况

                    max_col = ws.max_column
                    new_row = max_row + 1
                    
                    # 合并新行的所有列
                    ws.merge_cells(start_row=new_row, start_column=1, end_row=new_row, end_column=max_col)
                    
                    # 设置内容
                    target_cell = ws.cell(row=new_row, column=1)
                    target_cell.value = "应答人:安徽易德人力科技股份有限公司(盖单位章)\n日期:2025年12月18日"
                    
                    # 设置样式：居右显示，字体是宋体14号字
                    target_cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
                    target_cell.font = Font(name='宋体', size=14)
                    
                    # 调整行高以适应两行文本 (可选，视情况而定)
                    ws.row_dimensions[new_row].height = 40

                    # 3. 新建第二个sheet，sheet名是“盖章版” (空白sheet)
                    # 如果已经存在，先删除（虽然一般新建不会有）
                    if "盖章版" in wb.sheetnames:
                        del wb["盖章版"]
                    ws_stamp = wb.create_sheet("盖章版")

                    # 4. 删除除了“可编辑版本”和“盖章版”之外的其他sheet
                    for sheet_name in wb.sheetnames:
                        if sheet_name not in ["可编辑版本", "盖章版"]:
                            del wb[sheet_name]

                    # 5. 设置列宽
                    ws.column_dimensions['E'].width = 36
                    ws.column_dimensions['F'].width = 20

                    wb.save(file_path)
                    self.log(f"已处理Excel文件: {filename}")
                    
                except Exception as e:
                    self.log(f"处理Excel文件 {filename} 时出错: {e}")

class UnzipGUI(QMainWindow):
    """解压工具GUI类"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.worker = None
    
    def init_ui(self):
        """初始化用户界面"""
        # 设置窗口标题和大小
        self.setWindowTitle("国网电力招标文件解压缩程序")
        self.setGeometry(100, 100, 900, 600)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建目录选择区域
        dir_layout = QVBoxLayout()
        
        # 源目录选择
        source_layout = QHBoxLayout()
        source_label = QLabel("源目录:")
        source_label.setMinimumWidth(60)
        self.source_edit = QLineEdit()
        self.source_edit.setReadOnly(True)
        source_btn = QPushButton("浏览...")
        source_btn.clicked.connect(self.select_source_dir)
        
        source_layout.addWidget(source_label)
        source_layout.addWidget(self.source_edit)
        source_layout.addWidget(source_btn)
        
        # 目标目录选择
        target_layout = QHBoxLayout()
        target_label = QLabel("目标目录:")
        target_label.setMinimumWidth(60)
        self.target_edit = QLineEdit()
        self.target_edit.setReadOnly(True)
        target_btn = QPushButton("浏览...")
        target_btn.clicked.connect(self.select_target_dir)
        
        target_layout.addWidget(target_label)
        target_layout.addWidget(self.target_edit)
        target_layout.addWidget(target_btn)
        
        # 将源目录和目标目录布局添加到目录选择布局
        dir_layout.addLayout(source_layout)
        dir_layout.addLayout(target_layout)
        
        # 创建执行模式选择区域
        mode_group_box = QGroupBox("执行模式")
        mode_layout = QVBoxLayout()
        
        # 创建单选按钮组
        self.mode_button_group = QButtonGroup()
        
        # 创建5个单选按钮
        self.radio_full = QRadioButton("完整流程")
        self.radio_copy_engineering = QRadioButton("仅复制工程量清单")
        self.radio_convert_doc = QRadioButton("仅转换技术规范书")
        self.radio_check_technical = QRadioButton("仅检查技术规范书")
        self.radio_check_engineering = QRadioButton("仅检查工作量清单")
        
        # 添加到按钮组
        self.mode_button_group.addButton(self.radio_full, 0)
        self.mode_button_group.addButton(self.radio_copy_engineering, 1)
        self.mode_button_group.addButton(self.radio_convert_doc, 2)
        self.mode_button_group.addButton(self.radio_check_technical, 3)
        self.mode_button_group.addButton(self.radio_check_engineering, 4)
        
        # 默认选中"完整流程"
        self.radio_full.setChecked(True)
        
        # 添加到布局
        mode_layout.addWidget(self.radio_full)
        mode_layout.addWidget(self.radio_copy_engineering)
        mode_layout.addWidget(self.radio_convert_doc)
        mode_layout.addWidget(self.radio_check_technical)
        mode_layout.addWidget(self.radio_check_engineering)
        
        mode_group_box.setLayout(mode_layout)
        
        # 创建操作按钮区域
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始处理")
        self.start_btn.clicked.connect(self.start_processing)
        self.stop_btn = QPushButton("停止处理")
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setEnabled(False)
        self.open_folder_btn = QPushButton("打开目标文件夹")
        self.open_folder_btn.clicked.connect(self.open_target_folder)
        self.open_folder_btn.setEnabled(False)
        
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addWidget(self.open_folder_btn)
        
        # 创建状态和进度区域
        status_layout = QHBoxLayout()
        self.status_label = QLabel("准备就绪")
        self.status_label.setStyleSheet("color: blue;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # 不确定进度
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.progress_bar)
        
        # 创建日志区域
        log_label = QLabel("处理日志:")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setLineWrapMode(QTextEdit.WidgetWidth)
        self.log_text.setAcceptRichText(False)
        
        # 添加所有布局到主布局
        main_layout.addLayout(dir_layout)
        main_layout.addWidget(mode_group_box)
        main_layout.addLayout(btn_layout)
        main_layout.addLayout(status_layout)
        main_layout.addWidget(log_label)
        main_layout.addWidget(self.log_text)
        
        # 设置布局比例
        main_layout.setStretch(0, 1)  # dir_layout
        main_layout.setStretch(1, 1)  # mode_group_box
        main_layout.setStretch(2, 1)  # btn_layout
        main_layout.setStretch(3, 1)  # status_layout
        main_layout.setStretch(4, 1)  # log_label
        main_layout.setStretch(5, 10)  # log_text
    
    def select_source_dir(self):
        """选择源目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择源目录", "/")
        if dir_path:
            self.source_edit.setText(dir_path)
    
    def select_target_dir(self):
        """选择目标目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择目标目录")
        if dir_path:
            self.target_edit.setText(dir_path)
            # 重置打开文件夹按钮状态
            self.open_folder_btn.setEnabled(False)
    
    def start_processing(self):
        """开始处理"""
        # 获取源目录和目标目录
        source_dir = self.source_edit.text().strip()
        target_dir = self.target_edit.text().strip()
        
        # 获取选中的执行模式
        mode = 'full'  # 默认完整流程
        if self.radio_copy_engineering.isChecked():
            mode = 'copy_engineering'
        elif self.radio_convert_doc.isChecked():
            mode = 'convert_doc'
        elif self.radio_check_technical.isChecked():
            mode = 'check_technical'
        elif self.radio_check_engineering.isChecked():
            mode = 'check_engineering'
        
        # 验证目录
        # 只有完整流程才需要源目录，其他模式只需要目标目录
        if mode == 'full':
            if not source_dir:
                QMessageBox.warning(self, "警告", "请选择源目录")
                return
            
            if not target_dir:
                QMessageBox.warning(self, "警告", "请选择目标目录")
                return
            
            if source_dir == target_dir:
                QMessageBox.warning(self, "警告", "源目录和目标目录不能相同")
                return
            
            # 检查源目录是否存在
            if not os.path.exists(source_dir):
                QMessageBox.warning(self, "警告", f"源目录不存在: {source_dir}")
                return
        else:
            # 非完整流程模式，只需要目标目录
            if not target_dir:
                QMessageBox.warning(self, "警告", "请选择目标目录")
                return
            
            # 检查目标目录是否存在
            if not os.path.exists(target_dir):
                QMessageBox.warning(self, "警告", f"目标目录不存在: {target_dir}")
                return
        
        # 创建并启动工作线程
        self.worker = UnzipWorker(source_dir, target_dir, mode)
        self.worker.progress_updated.connect(self.update_log)
        self.worker.status_changed.connect(self.update_status)
        self.worker.finished.connect(self.process_finished)
        self.worker.error.connect(self.show_error)
        
        # 更新UI状态
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("开始处理...")
        self.status_label.setStyleSheet("color: green;")
        self.progress_bar.setVisible(True)
        self.log_text.clear()
        
        # 启动线程
        self.worker.start()
    
    def stop_processing(self):
        """停止处理"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.status_label.setText("正在停止...")
            self.status_label.setStyleSheet("color: orange;")
            self.stop_btn.setEnabled(False)
    
    def update_log(self, message):
        """更新日志显示"""
        self.log_text.append(message)
        # 自动滚动到底部
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log_text.setTextCursor(cursor)
    
    def update_status(self, status):
        """更新状态显示"""
        self.status_label.setText(status)
    
    def process_finished(self):
        """处理完成"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.open_folder_btn.setEnabled(True)
        self.status_label.setText("处理完成")
        self.status_label.setStyleSheet("color: blue;")
        self.progress_bar.setVisible(False)
        
        # 显示完成消息
        QMessageBox.information(self, "完成", "文件处理已完成！")
    
    def show_error(self, error_message):
        """显示错误消息"""
        self.status_label.setText("处理出错")
        self.status_label.setStyleSheet("color: red;")
        self.update_log(f"错误: {error_message}")
        QMessageBox.critical(self, "错误", error_message)
    
    def open_target_folder(self):
        """打开目标文件夹"""
        target_dir = self.target_edit.text().strip()
        if not target_dir:
            QMessageBox.warning(self, "警告", "请先选择目标目录")
            return
        
        if not os.path.exists(target_dir):
            QMessageBox.warning(self, "警告", f"目标目录不存在: {target_dir}")
            return
        
        try:
            # 根据操作系统选择合适的命令
            system = platform.system()
            if system == 'Windows':
                os.startfile(target_dir)
            elif system == 'Darwin':  # macOS
                subprocess.run(['open', target_dir])
            elif system == 'Linux':
                subprocess.run(['xdg-open', target_dir])
            else:
                # 其他系统尝试使用xdg-open
                try:
                    subprocess.run(['xdg-open', target_dir])
                except:
                    QMessageBox.warning(self, "警告", "不支持的操作系统，无法自动打开文件夹")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开文件夹时出错: {str(e)}")
    
    def closeEvent(self, event):
        """窗口关闭事件"""
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, "确认关闭", 
                "处理正在进行中，确定要关闭吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.No:
                event.ignore()
                return
            else:
                self.worker.stop()
                self.worker.wait(1000)  # 等待1秒让线程停止
        event.accept()

def main():
    """主函数"""
    # 确保中文显示正常
    import os
    os.environ['QT_FONT_DPI'] = '96'  # 设置字体DPI，避免字体模糊
    
    app = QApplication(sys.argv)
    
    # 设置应用样式
    app.setStyle('Fusion')
    
    # 创建并显示主窗口
    window = UnzipGUI()
    window.show()
    
    # 运行应用
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
