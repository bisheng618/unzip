#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从log.txt中提取成功删除目录节点的日志条目
"""

import re


def extract_success_logs(log_file_path, output_file_path=None):
    """
    从日志文件中提取包含"[成功] 已删除第 2 个目录节点"的完整日志块
    
    Args:
        log_file_path: 日志文件路径
        output_file_path: 输出文件路径,如果为None则只打印到控制台
    
    Returns:
        提取的日志条目列表
    """
    try:
        with open(log_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print(f"错误: 找不到文件 {log_file_path}")
        return []
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return []
    
    # 按照"[匹配成功]"分割日志块
    blocks = re.split(r'(\[匹配成功\])', content)
    
    success_entries = []
    
    # 重新组合被分割的块
    for i in range(1, len(blocks), 2):
        if i + 1 < len(blocks):
            block = blocks[i] + blocks[i + 1]
            
            # 检查是否包含成功删除的标记
            if '[成功] 已删除第 2 个目录节点' in block:
                success_entries.append(block.strip())
    
    # 输出结果
    if success_entries:
        print(f"共提取到 {len(success_entries)} 条成功删除的日志条目:\n")
        print("=" * 80)
        
        for idx, entry in enumerate(success_entries, 1):
            print(f"\n条目 {idx}:")
            print("-" * 80)
            print(entry)
            print("-" * 80)
        
        # 如果指定了输出文件,则写入文件
        if output_file_path:
            try:
                with open(output_file_path, 'w', encoding='utf-8') as f:
                    for idx, entry in enumerate(success_entries, 1):
                        f.write(f"条目 {idx}:\n")
                        f.write("=" * 80 + "\n")
                        f.write(entry + "\n")
                        f.write("=" * 80 + "\n\n")
                print(f"\n结果已保存到: {output_file_path}")
            except Exception as e:
                print(f"\n写入输出文件时出错: {e}")
    else:
        print("未找到符合条件的日志条目")
    
    return success_entries


if __name__ == "__main__":
    # 日志文件路径
    log_file = "log.txt"
    
    # 输出文件路径(可选)
    output_file = "success_logs.txt"
    
    # 提取日志
    extract_success_logs(log_file, output_file)
