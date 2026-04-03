#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试打开目标文件夹功能
"""

import os
import sys
import platform
import subprocess

def open_folder(folder_path):
    """打开文件夹"""
    if not os.path.exists(folder_path):
        print(f"文件夹不存在: {folder_path}")
        return False
    
    try:
        # 根据操作系统选择合适的命令
        system = platform.system()
        if system == 'Windows':
            os.startfile(folder_path)
        elif system == 'Darwin':  # macOS
            subprocess.run(['open', folder_path])
        elif system == 'Linux':
            subprocess.run(['xdg-open', folder_path])
        else:
            # 其他系统尝试使用xdg-open
            try:
                subprocess.run(['xdg-open', folder_path])
            except:
                print("不支持的操作系统，无法自动打开文件夹")
                return False
        
        print(f"已打开文件夹: {folder_path}")
        return True
    except Exception as e:
        print(f"打开文件夹时出错: {str(e)}")
        return False

if __name__ == "__main__":
    # 测试打开当前目录
    current_dir = os.getcwd()
    print(f"当前操作系统: {platform.system()}")
    print(f"测试打开文件夹: {current_dir}")
    open_folder(current_dir)