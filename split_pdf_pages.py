#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF拆分工具
将"拟委任的主要人员汇总表.pdf"拆分为两个文件：
1. 前两页
2. 剩余页
"""

import os
from PyPDF2 import PdfReader, PdfWriter

def split_pdf():
    source_dir = "/Users/bisheng/Downloads/lulili2/"
    source_file = os.path.join(source_dir, "拟委任的主要人员汇总表.pdf")
    
    if not os.path.exists(source_file):
        print(f"错误: 文件不存在: {source_file}")
        return

    try:
        reader = PdfReader(source_file)
        total_pages = len(reader.pages)
        print(f"源文件共有 {total_pages} 页")

        if total_pages < 2:
            print("警告: 文件少于2页，无法按要求拆分")
            return

        # 创建两个Writer对象
        writer1 = PdfWriter()
        writer2 = PdfWriter()

        # 添加页面
        # Part 1: 前2页 (索引 0, 1)
        for i in range(min(2, total_pages)):
            writer1.add_page(reader.pages[i])

        # Part 2: 剩余页 (索引 2 到最后)
        for i in range(2, total_pages):
            writer2.add_page(reader.pages[i])

        # 保存文件
        output1 = os.path.join(source_dir, "拟委任的主要人员汇总表_Part1.pdf")
        output2 = os.path.join(source_dir, "拟委任的主要人员汇总表_Part2.pdf")

        with open(output1, "wb") as f1:
            writer1.write(f1)
        print(f"已生成: {output1} (共 {len(writer1.pages)} 页)")

        if len(writer2.pages) > 0:
            with open(output2, "wb") as f2:
                writer2.write(f2)
            print(f"已生成: {output2} (共 {len(writer2.pages)} 页)")
        else:
            print("没有剩余页面，未生成Part2")

    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == "__main__":
    split_pdf()
