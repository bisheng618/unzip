#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据 project.txt 过滤并转换 Excel 为 PDF
脚本会读取 project.txt 中的每一行作为项目名称关键词。
搜索当前目录下文件名包含“完成报价”且包含项目名称关键词的 Excel 文件。
将这些 Excel 文件的第一个工作表转换为 PDF。
最后将生成的 PDF 合并为一个总文件。
"""

import os
import subprocess
import tempfile
import shutil
from pathlib import Path
from PyPDF2 import PdfMerger
import openpyxl
from openpyxl.worksheet.page import PageMargins

def find_soffice():
    """查找 LibreOffice 的 soffice 可执行文件"""
    possible_paths = [
        '/opt/homebrew/bin/soffice',
        '/usr/local/bin/soffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        shutil.which('soffice'),
        shutil.which('libreoffice'),
    ]
    for path in possible_paths:
        if path and os.path.exists(path):
            return path
    return None

def read_project_names(txt_path):
    """从 project.txt 读取项目名称列表"""
    if not os.path.exists(txt_path):
        print(f"错误: 找不到文件 {txt_path}")
        return []
    with open(txt_path, 'r', encoding='utf-8') as f:
        projects = [line.strip() for line in f if line.strip()]
    return projects

def find_matching_excel_files(root_dir, project_names):
    """递归查找匹配的 Excel 文件，并记录匹配到的项目名称"""
    matched_files = []
    matched_project_names = set()
    print(f"开始在 {root_dir} 中搜索匹配的 Excel 文件...")
    
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.startswith('~$'): continue # 跳过 Excel 临时文件
            if not (file.endswith('.xlsx') or file.endswith('.xls')): continue
            
            # 条件 1：文件名包含“完成报价”
            if "完成报价" not in file:
                continue
            
            # 条件 2：文件名包含 project.txt 中的任意一个项目名
            for proj in project_names:
                if proj in file:
                    full_path = os.path.join(root, file)
                    matched_files.append(full_path)
                    matched_project_names.add(proj)
                    print(f"  找到匹配文件: {file}")
                    # 注意：这里不 break 是为了防止一个文件匹配多个项目关键词，
                    # 也可以根据需求调整为 break 如果一个文件只代表一个项目
                
    return sorted(list(set(matched_files))), matched_project_names

def create_temp_excel_with_first_sheet(excel_path, temp_dir):
    """创建只包含第一个 sheet 的临时 Excel 文件，并设置打印布局"""
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        
        new_wb = openpyxl.Workbook()
        new_wb.remove(new_wb.active)
        
        source_sheet = wb.worksheets[0]
        sheet_name = source_sheet.title
        target_sheet = new_wb.create_sheet(sheet_name)
        
        # 复制内容和样式
        for row in source_sheet.iter_rows():
            for cell in row:
                target_cell = target_sheet[cell.coordinate]
                target_cell.value = cell.value
                if cell.has_style:
                    target_cell.font = cell.font.copy()
                    target_cell.border = cell.border.copy()
                    target_cell.fill = cell.fill.copy()
                    target_cell.number_format = cell.number_format
                    target_cell.protection = cell.protection.copy()
                    target_cell.alignment = cell.alignment.copy()
        
        # 复制行列维度
        for col_letter, col_dim in source_sheet.column_dimensions.items():
            target_sheet.column_dimensions[col_letter].width = col_dim.width
        for row_num, row_dim in source_sheet.row_dimensions.items():
            target_sheet.row_dimensions[row_num].height = row_dim.height
        for merged_cell in source_sheet.merged_cells.ranges:
            target_sheet.merge_cells(str(merged_cell))
            
        # 打印设置：适应页面宽度
        target_sheet.page_setup.orientation = 'portrait'
        target_sheet.page_setup.paperSize = 9 # A4
        target_sheet.page_setup.fitToPage = True
        target_sheet.page_setup.fitToWidth = 1
        target_sheet.page_setup.fitToHeight = 0
        target_sheet.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
        
        temp_excel_path = os.path.join(temp_dir, f"temp_{os.path.basename(excel_path)}")
        new_wb.save(temp_excel_path)
        wb.close()
        new_wb.close()
        return temp_excel_path
    except Exception as e:
        print(f"  错误: 处理 Excel {os.path.basename(excel_path)} 时出错: {e}")
        return None

def convert_to_pdf(excel_path, output_dir, soffice_path):
    """调用 LibreOffice 转换为 PDF"""
    try:
        cmd = [soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, excel_path]
        subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        pdf_name = Path(excel_path).stem + '.pdf'
        pdf_path = os.path.join(output_dir, pdf_name)
        return pdf_path if os.path.exists(pdf_path) else None
    except Exception as e:
        print(f"  转换 PDF 出错: {e}")
        return None

def main():
    soffice_path = find_soffice()
    if not soffice_path:
        print("错误: 未找到 LibreOffice。请安装后再试。")
        return

    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_txt = os.path.join(script_dir, "project.txt")
    
    project_names = read_project_names(project_txt)
    if not project_names:
        print("未在 project.txt 中找到任何项目名称。")
        return
    
    print(f"已读取 {len(project_names)} 个项目名称。")
    
    # 允许用户输入搜索目录
    source_directory = input("\n请输入要搜索文件的根目录路径: ").strip()
    if not source_directory or not os.path.exists(source_directory):
        print(f"错误: 路径不存在或无效: {source_directory}")
        return

    excel_files, matched_project_names = find_matching_excel_files(source_directory, project_names)
    
    # 打印未找到的项目
    unmatched = [p for p in project_names if p not in matched_project_names]
    if unmatched:
        print("\n" + "!" * 10 + " 未找到以下项目对应的“完成报价”文件 " + "!" * 10)
        for name in unmatched:
            print(f"  [未找到] {name}")
        print("!" * 54 + "\n")

    if not excel_files:
        print("未找到符合条件的 Excel 文件，程序结束。")
        return
    
    print(f"\n找到 {len(excel_files)} 个匹配的 Excel 文件，开始转换...")
    
    temp_dir = tempfile.mkdtemp()
    pdf_files = []
    
    try:
        for i, excel in enumerate(excel_files, 1):
            print(f"[{i}/{len(excel_files)}] 处理: {os.path.basename(excel)}")
            temp_excel = create_temp_excel_with_first_sheet(excel, temp_dir)
            if not temp_excel: continue
            
            pdf_path = convert_to_pdf(temp_excel, temp_dir, soffice_path)
            if pdf_path:
                pdf_files.append(pdf_path)
                print(f"  ✓ 成功转换")
        
        if pdf_files:
            output_pdf = os.path.join(source_directory, "项目报价汇总表_合并.pdf")
            merger = PdfMerger()
            for pdf in pdf_files:
                merger.append(pdf)
            merger.write(output_pdf)
            merger.close()
            print(f"\n" + "=" * 40)
            print(f"✓ 任务完成！")
            print(f"✓ 已合并 {len(pdf_files)} 个 PDF")
            print(f"✓ 输出文件: {output_pdf}")
            print(f"========================================")
        else:
            print("\n没有文件被成功转换成 PDF。")
            
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    main()
