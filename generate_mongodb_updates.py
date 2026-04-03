#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从Excel文件生成MongoDB更新语句的程序

此程序读取Excel文件中的contractcode和idno列，
并为每个记录生成相应的MongoDB更新语句。
"""

import pandas as pd


def generate_mongodb_updates(excel_path, output_path):
    """
    从Excel文件生成MongoDB更新语句
    
    Args:
        excel_path (str): Excel文件路径
        output_path (str): 输出文件路径
    """
    print(f"正在读取Excel文件: {excel_path}")
    
    # 读取Excel文件
    df = pd.read_excel(excel_path)
    
    # 检查必要的列是否存在
    required_columns = ['contractcode', 'idno']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Excel文件缺少必需的列: {col}")
    
    print(f"找到 {len(df)} 条记录")
    
    # 生成更新语句
    update_statements = []
    for index, row in df.iterrows():
        contractcode = row['contractcode']
        idno = row['idno']
        
        # 格式化更新语句
        statement = f'db.getCollection("actual_serviceitem_{contractcode}").updateMany(\n' \
                   f'  {{"idno":"{idno}","categoryid":"house"\n' \
                   f'  }},\n' \
                   f'  {{\n' \
                   f'    $set: {{\n' \
                   f'      del: true\n' \
                   f'    }}\n' \
                   f'  }}\n' \
                   f');'
        update_statements.append(statement)
    
    # 写入输出文件
    with open(output_path, 'w', encoding='utf-8') as f:
        for i, statement in enumerate(update_statements):
            f.write(statement)
            if i < len(update_statements) - 1:  # 不在最后一个语句后面加空行
                f.write('\n\n')
    
    print(f"成功生成 {len(update_statements)} 条MongoDB更新语句")
    print(f"输出文件: {output_path}")


def main():
    # 输入和输出文件路径
    excel_path = "/Users/bisheng/Downloads/prd_hro_files_db_tb_files_personnel_3.xlsx"
    output_path = "/Users/bisheng/PycharmProjects/unzipfile/mongodb_updates_3.js"
    
    try:
        generate_mongodb_updates(excel_path, output_path)
        print("程序执行完成!")
    except FileNotFoundError:
        print(f"错误: 找不到Excel文件 {excel_path}")
    except ValueError as e:
        print(f"错误: {e}")
    except Exception as e:
        print(f"发生未知错误: {e}")


if __name__ == "__main__":
    main()