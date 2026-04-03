#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import argparse
import logging
from datetime import datetime
from pathlib import Path

def setup_logging():
    """设置日志记录"""
    log_filename = f"delete_auth_files_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def find_authorized_documents(directory_path):
    """
    递归查找指定目录及子目录下所有文件名含有"+授权委托书"的docx文件
    
    Args:
        directory_path (str): 要搜索的目录路径
    
    Returns:
        list: 匹配的文件路径列表
    """
    matching_files = []
    
    # 检查目录是否存在
    if not os.path.exists(directory_path):
        return matching_files
    
    if not os.path.isdir(directory_path):
        return matching_files
    
    # 递归遍历目录
    try:
        path = Path(directory_path)
        for file_path in path.rglob("*+授权委托书*.docx"):
            if file_path.is_file() and file_path.suffix.lower() == '.docx':
                matching_files.append(str(file_path))
    except Exception as e:
        print(f"遍历目录时出错: {e}")
    
    return matching_files

def delete_authorized_documents(directory_path, dry_run=False, logger=None):
    """
    删除指定目录及子目录下所有文件名含有"+授权委托书"的docx文件
    
    Args:
        directory_path (str): 要搜索的目录路径
        dry_run (bool): 是否为试运行模式（仅显示将要删除的文件）
        logger (Logger): 日志记录器
    """
    if logger is None:
        logger = logging.getLogger(__name__)
    
    # 检查目录是否存在
    if not os.path.exists(directory_path):
        error_msg = f"错误: 目录 '{directory_path}' 不存在"
        print(error_msg)
        logger.error(error_msg)
        return
    
    if not os.path.isdir(directory_path):
        error_msg = f"错误: '{directory_path}' 不是一个有效的目录"
        print(error_msg)
        logger.error(error_msg)
        return
    
    # 查找匹配的文件
    files_to_delete = find_authorized_documents(directory_path)
    
    if not files_to_delete:
        info_msg = "未找到包含'+授权委托书'的docx文件"
        print(info_msg)
        logger.info(info_msg)
        return
    
    info_msg = f"找到 {len(files_to_delete)} 个需要删除的文件:"
    print(info_msg)
    logger.info(info_msg)
    
    for file_path in files_to_delete:
        filename = os.path.basename(file_path)
        relative_path = os.path.relpath(file_path, directory_path)
        print(f"  - {relative_path}")
        logger.info(f"匹配文件: {relative_path}")
    
    if dry_run:
        info_msg = "\n试运行模式: 以上文件将不会被实际删除"
        print(info_msg)
        logger.info(info_msg)
        return
    
    # 确认删除
    confirm = input("\n确认删除以上文件吗? (y/N): ")
    if confirm.lower() in ['y', 'yes']:
        deleted_count = 0
        for file_path in files_to_delete:
            try:
                os.remove(file_path)
                success_msg = f"已删除: {os.path.relpath(file_path, directory_path)}"
                print(success_msg)
                logger.info(success_msg)
                deleted_count += 1
            except Exception as e:
                error_msg = f"删除失败 {os.path.relpath(file_path, directory_path)}: {e}"
                print(error_msg)
                logger.error(error_msg)
        
        summary_msg = f"\n删除完成，共删除 {deleted_count} 个文件"
        print(summary_msg)
        logger.info(summary_msg)
    else:
        cancel_msg = "取消删除操作"
        print(cancel_msg)
        logger.info(cancel_msg)

def main():
    parser = argparse.ArgumentParser(description="删除包含'+授权委托书'的docx文件")
    parser.add_argument("--directory", "-d", default="/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/国网安徽电力第一片区2025年第十次服务区域联合授权竞争性谈判采购-2025-12-02-AM1：00-ECP/",
                        help="要搜索的目录路径")
    parser.add_argument("--dry-run", action="store_true", 
                        help="试运行模式：只显示将要删除的文件，不实际删除")
    parser.add_argument("--log", action="store_true", 
                        help="启用详细日志记录")
    
    args = parser.parse_args()
    
    # 设置日志
    logger = setup_logging()
    
    # 指定要处理的目录
    target_directory = args.directory
    
    print("删除包含'+授权委托书'的docx文件工具")
    print("=" * 50)
    print(f"目标目录: {target_directory}")
    if args.dry_run:
        print("运行模式: 试运行模式")
    print()
    
    delete_authorized_documents(target_directory, args.dry_run, logger)

if __name__ == "__main__":
    main()
