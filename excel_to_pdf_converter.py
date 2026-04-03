#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel转PDF工具
使用LibreOffice将Excel文件的"可编辑版"sheet转换为PDF，并合并成一个PDF文件
"""

import os
import glob
import subprocess
import tempfile
import shutil
from pathlib import Path
from PyPDF2 import PdfMerger


def find_soffice():
    """
    查找LibreOffice的soffice可执行文件
    
    Returns:
        str: soffice路径，如果未找到则返回None
    """
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


def find_excel_files(directory):
    """
    查找指定目录下的所有Excel文件，且文件名包含“完成报价”
    
    Args:
        directory: 目标目录路径
        
    Returns:
        Excel文件路径列表
    """
    excel_files = []
    patterns = ['*.xlsx', '*.xls', '*.xlsm']
    
    for pattern in patterns:
        files = glob.glob(os.path.join(directory, pattern))
        # 仅保留文件名包含 "完成报价" 的文件
        excel_files.extend([f for f in files if "完成报价" in os.path.basename(f)])
    
    return sorted(excel_files)


def create_temp_excel_with_sheet(excel_path, sheet_name, temp_dir):
    """
    创建只包含指定sheet的临时Excel文件，并设置页面打印属性
    
    Args:
        excel_path: 原始Excel文件路径
        sheet_name: 要保留的工作表名称
        temp_dir: 临时目录
        
    Returns:
        str: 临时Excel文件路径，如果失败返回None
    """
    try:
        import openpyxl
        from openpyxl.worksheet.page import PageMargins, PrintPageSetup
        
        # 打开原始Excel文件
        wb = openpyxl.load_workbook(excel_path)
        
        # 检查sheet是否存在
        if sheet_name not in wb.sheetnames:
            print(f"  警告: 文件 {os.path.basename(excel_path)} 中未找到工作表 '{sheet_name}'")
            wb.close()
            return None
        
        # 创建新工作簿，只包含目标sheet
        new_wb = openpyxl.Workbook()
        new_wb.remove(new_wb.active)  # 删除默认sheet
        
        # 复制目标sheet
        source_sheet = wb[sheet_name]
        target_sheet = new_wb.create_sheet(sheet_name)
        
        # 复制所有单元格（包括样式）
        for row in source_sheet.iter_rows():
            for cell in row:
                target_cell = target_sheet[cell.coordinate]
                target_cell.value = cell.value
                
                # 复制样式
                if cell.has_style:
                    target_cell.font = cell.font.copy()
                    target_cell.border = cell.border.copy()
                    target_cell.fill = cell.fill.copy()
                    target_cell.number_format = cell.number_format
                    target_cell.protection = cell.protection.copy()
                    target_cell.alignment = cell.alignment.copy()
        
        # 复制列宽
        for col_letter, col_dim in source_sheet.column_dimensions.items():
            target_sheet.column_dimensions[col_letter].width = col_dim.width
        
        # 复制行高
        for row_num, row_dim in source_sheet.row_dimensions.items():
            target_sheet.row_dimensions[row_num].height = row_dim.height
        
        # 复制合并单元格
        for merged_cell in source_sheet.merged_cells.ranges:
            target_sheet.merge_cells(str(merged_cell))
        
        # 设置页面打印属性 - 关键：让所有列适应在一页宽度内
        target_sheet.page_setup.orientation = 'portrait'  # 竖向
        target_sheet.page_setup.paperSize = 9  # A4纸张
        target_sheet.page_setup.fitToPage = True  # 启用适应页面
        target_sheet.page_setup.fitToWidth = 1  # 所有列适应在1页宽度内
        target_sheet.page_setup.fitToHeight = 0  # 高度不限制（0表示不限制）
        
        # 设置较小的页边距
        target_sheet.page_margins = PageMargins(
            left=0.5, right=0.5, top=0.5, bottom=0.5,
            header=0.3, footer=0.3
        )
        
        # 保存临时文件
        temp_excel_path = os.path.join(temp_dir, f"temp_{os.path.basename(excel_path)}")
        new_wb.save(temp_excel_path)
        
        wb.close()
        new_wb.close()
        
        return temp_excel_path
        
    except Exception as e:
        print(f"  错误: 创建临时Excel文件时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


def convert_excel_to_pdf_libreoffice(excel_path, output_dir, soffice_path):
    """
    使用LibreOffice将Excel文件转换为PDF
    
    Args:
        excel_path: Excel文件路径
        output_dir: 输出目录
        soffice_path: LibreOffice soffice可执行文件路径
        
    Returns:
        str: 生成的PDF文件路径，如果失败返回None
    """
    try:
        # LibreOffice转换命令
        # --headless: 无界面模式
        # --convert-to pdf: 转换为PDF
        # --outdir: 输出目录
        cmd = [
            soffice_path,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            excel_path
        ]
        
        # 执行转换
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode != 0:
            print(f"  错误: LibreOffice转换失败: {result.stderr}")
            return None
        
        # 生成的PDF文件名
        pdf_filename = Path(excel_path).stem + '.pdf'
        pdf_path = os.path.join(output_dir, pdf_filename)
        
        if os.path.exists(pdf_path):
            return pdf_path
        else:
            print(f"  错误: PDF文件未生成: {pdf_path}")
            return None
            
    except subprocess.TimeoutExpired:
        print(f"  错误: 转换超时")
        return None
    except Exception as e:
        print(f"  错误: 转换过程出错: {str(e)}")
        return None


def merge_pdfs(pdf_files, output_path):
    """
    合并多个PDF文件为一个
    
    Args:
        pdf_files: PDF文件路径列表
        output_path: 输出PDF路径
    """
    merger = PdfMerger()
    
    try:
        for pdf_file in pdf_files:
            if os.path.exists(pdf_file):
                merger.append(pdf_file)
                print(f"  添加到合并列表: {os.path.basename(pdf_file)}")
        
        merger.write(output_path)
        print(f"\n成功合并PDF: {output_path}")
        
    except Exception as e:
        print(f"合并PDF时出错: {str(e)}")
        
    finally:
        merger.close()


def main():
    """主函数"""
    # 查找LibreOffice
    soffice_path = find_soffice()
    if not soffice_path:
        print("错误: 未找到LibreOffice，请先安装LibreOffice")
        print("可以通过以下方式安装:")
        print("  brew install --cask libreoffice")
        return
    
    print(f"使用LibreOffice: {soffice_path}")
    print("-" * 60)
    
    # 目标文件夹路径
    source_directory = input("请输入要处理的文件夹路径: ").strip()
    if not source_directory or not os.path.exists(source_directory):
        print(f"错误: 路径不存在: {source_directory}")
        return
    
    # 输出PDF文件路径
    output_directory = source_directory
    final_pdf_name = "工程量清单_汇总_合并.pdf"
    final_pdf_path = os.path.join(output_directory, final_pdf_name)
    
    # 工作表名称
    sheet_name = "可编辑版本"
    
    print(f"开始处理文件夹: {source_directory}")
    print(f"目标工作表: {sheet_name}")
    print("-" * 60)
    
    # 查找所有Excel文件
    excel_files = find_excel_files(source_directory)
    
    if not excel_files:
        print("未找到任何Excel文件！")
        return
    
    print(f"找到 {len(excel_files)} 个Excel文件\n")
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    pdf_output_dir = tempfile.mkdtemp()
    converted_pdfs = []
    
    try:
        # 逐个转换Excel文件
        for idx, excel_file in enumerate(excel_files, 1):
            print(f"[{idx}/{len(excel_files)}] 处理: {os.path.basename(excel_file)}")
            
            # 创建只包含目标sheet的临时Excel文件
            temp_excel = create_temp_excel_with_sheet(excel_file, sheet_name, temp_dir)
            
            if not temp_excel:
                continue
            
            # 使用LibreOffice转换为PDF
            pdf_path = convert_excel_to_pdf_libreoffice(temp_excel, pdf_output_dir, soffice_path)
            
            if pdf_path:
                # 重命名PDF文件以包含原始文件名
                final_pdf_name_single = f"{Path(excel_file).stem}_{sheet_name}.pdf"
                final_pdf_path_single = os.path.join(pdf_output_dir, final_pdf_name_single)
                
                if os.path.exists(pdf_path):
                    os.rename(pdf_path, final_pdf_path_single)
                    converted_pdfs.append(final_pdf_path_single)
                    print(f"  ✓ 成功转换: {os.path.basename(excel_file)}")
        
        print("\n" + "=" * 60)
        
        # 合并所有PDF
        if converted_pdfs:
            print(f"开始合并 {len(converted_pdfs)} 个PDF文件...")
            merge_pdfs(converted_pdfs, final_pdf_path)
            print("=" * 60)
            print(f"\n✓ 所有操作完成！")
            print(f"✓ 最终PDF文件: {final_pdf_path}")
            print(f"✓ 成功转换: {len(converted_pdfs)}/{len(excel_files)} 个文件")
        else:
            print("没有成功转换的PDF文件，无法合并。")
    
    finally:
        # 清理临时文件
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        if os.path.exists(pdf_output_dir):
            shutil.rmtree(pdf_output_dir)
        print(f"\n已清理临时文件")


if __name__ == "__main__":
    main()
