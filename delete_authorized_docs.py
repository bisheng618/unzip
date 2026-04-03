#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import argparse
from pathlib import Path

def find_and_delete_authorized_docs(directory_path, dry_run=False):
    """
    递归查找并删除指定目录及子目录下所有文件名含有"+授权委托书"的docx文件
    
    Args:
        directory_path (str): 要搜索的目录路径
        dry_run (bool): 是否为试运行模式（仅显示将要删除的文件，不实际删除）
    """
    # 检查目录是否存在
    if not os.path.exists(directory_path):
        print(f"错误: 目录 '{directory_path}' 不存在")
        return
    
    if not os.path.isdir(directory_path):
        print(f"错误: '{directory_path}' 不是一个有效的目录")
        return
    
    # 查找匹配的文件
    matching_files = []
    try:
        path = Path(directory_path)
        for file_path in path.rglob("*+授权委托书*.docx"):
            if file_path.is_file() and file_path.suffix.lower() == '.docx':
                matching_files.append(file_path)
    except Exception as e:
        print(f"遍历目录时出错: {e}")
        return
    
    if not matching_files:
        print("未找到包含'+授权委托书'的docx文件")
        return
    
    print(f"找到 {len(matching_files)} 个需要删除的文件:")
    for file_path in matching_files:
        relative_path = file_path.relative_to(directory_path)
        print(f"  - {relative_path}")
    
    if dry_run:
        print("\n试运行模式: 以上文件将不会被实际删除")
        return
    
    # 确认删除
    confirm = input("\n确认删除以上文件吗? (y/N): ")
    if confirm.lower() in ['y', 'yes']:
        deleted_count = 0
        for file_path in matching_files:
            try:
                file_path.unlink()  # 删除文件
                relative_path = file_path.relative_to(directory_path)
                print(f"已删除: {relative_path}")
                deleted_count += 1
            except Exception as e:
                relative_path = file_path.relative_to(directory_path)
                print(f"删除失败 {relative_path}: {e}")
        
        print(f"\n删除完成，共删除 {deleted_count} 个文件")
    else:
        print("取消删除操作")

def main():
    parser = argparse.ArgumentParser(description="删除包含'+授权委托书'的docx文件")
    parser.add_argument("--directory", "-d", required=True,
                        help="要搜索的目录路径")
    parser.add_argument("--dry-run", action="store_true", 
                        help="试运行模式：只显示将要删除的文件，不实际删除")
    
    args = parser.parse_args()
    
    print("删除包含'+授权委托书'的docx文件工具")
    print("=" * 50)
    print(f"目标目录: {args.directory}")
    if args.dry_run:
        print("运行模式: 试运行模式")
    print()
    
    find_and_delete_authorized_docs(args.directory, args.dry_run)

if __name__ == "__main__":
    main()
