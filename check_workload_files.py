#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
扫描指定目录下的"上传版"文件夹,检查"工作量清单"相关文件的修改时间是否一致
"""

import os
from pathlib import Path
from datetime import datetime


def find_workload_files(upload_dir):
    """
    在指定目录中查找包含"工作量清单"的docx和xlsx文件
    
    Args:
        upload_dir: 上传版文件夹路径
        
    Returns:
        tuple: (docx_file, xlsx_file, docx_mtime, xlsx_mtime)
    """
    docx_file = None
    xlsx_file = None
    docx_mtime = None
    xlsx_mtime = None
    
    try:
        for file in os.listdir(upload_dir):
            file_path = os.path.join(upload_dir, file)
            
            # 跳过目录
            if os.path.isdir(file_path):
                continue
                
            # 检查是否包含"工作量清单"
            if "工作量清单" in file:
                if file.endswith('.docx'):
                    docx_file = file
                    docx_mtime = os.path.getmtime(file_path)
                elif file.endswith('.xlsx'):
                    xlsx_file = file
                    xlsx_mtime = os.path.getmtime(file_path)
    except Exception as e:
        print(f"读取目录 {upload_dir} 时出错: {e}")
    
    return docx_file, xlsx_file, docx_mtime, xlsx_mtime


def scan_directories(base_path):
    """
    扫描基础路径下的所有文件夹,查找"上传版"子文件夹并检查文件修改时间
    
    Args:
        base_path: 基础扫描路径
    """
    if not os.path.exists(base_path):
        print(f"错误: 路径不存在 - {base_path}")
        return
    
    print(f"开始扫描路径: {base_path}\n")
    print("=" * 80)
    
    inconsistent_folders = []
    
    # 遍历基础路径下的所有子目录
    for root, dirs, files in os.walk(base_path):
        # 检查当前目录是否包含"上传版"子文件夹
        if "上传版" in dirs:
            upload_dir = os.path.join(root, "上传版")
            
            # 查找工作量清单文件
            docx_file, xlsx_file, docx_mtime, xlsx_mtime = find_workload_files(upload_dir)
            
            # 如果同时找到了docx和xlsx文件
            if docx_file and xlsx_file:
                # 比较修改时间(精确到秒)
                if int(docx_mtime) != int(xlsx_mtime):
                    inconsistent_folders.append({
                        'folder': upload_dir,
                        'docx_file': docx_file,
                        'xlsx_file': xlsx_file,
                        'docx_mtime': datetime.fromtimestamp(docx_mtime),
                        'xlsx_mtime': datetime.fromtimestamp(xlsx_mtime)
                    })
    
    # 输出结果
    if inconsistent_folders:
        print(f"\n发现 {len(inconsistent_folders)} 个文件夹中的工作量清单文件修改时间不一致:\n")
        
        for idx, item in enumerate(inconsistent_folders, 1):
            print(f"{idx}. 文件夹: {item['folder']}")
            print(f"   DOCX文件: {item['docx_file']}")
            print(f"   修改时间: {item['docx_mtime'].strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"   XLSX文件: {item['xlsx_file']}")
            print(f"   修改时间: {item['xlsx_mtime'].strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"   时间差: {abs((item['docx_mtime'] - item['xlsx_mtime']).total_seconds())} 秒")
            print("-" * 80)
    else:
        print("\n未发现修改时间不一致的文件夹。")
    
    print("\n扫描完成!")


def main():
    # 指定要扫描的基础路径
    base_path = "/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/国网安徽电力第二片区2025年第十一次服务区域联合授权竞争性谈判采购_采购文件包"
    
    scan_directories(base_path)


if __name__ == "__main__":
    main()
