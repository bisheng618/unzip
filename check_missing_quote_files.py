import os

def scan_missing_quote_files(root_path):
    """
    扫描 root_path 下所有“上传版”子文件夹，
    检查是否缺少包含“完成报价”的 .docx 或 .xlsx 文档。
    """
    missing_docx = []
    missing_xlsx = []
    
    if not os.path.exists(root_path):
        print(f"错误：路径不存在 -> {root_path}")
        return

    print(f"开始扫描目录: {root_path}")
    print("-" * 50)

    for root, dirs, files in os.walk(root_path):
        # 只关注名为“上传版”的文件夹
        if os.path.basename(root) == "上传版":
            has_quote_docx = False
            has_quote_xlsx = False
            
            for file in files:
                lower_file = file.lower()
                if "完成报价" in file:
                    if lower_file.endswith(".docx"):
                        has_quote_docx = True
                    elif lower_file.endswith(".xlsx"):
                        has_quote_xlsx = True
            
            if not has_quote_docx:
                missing_docx.append(root)
            if not has_quote_xlsx:
                missing_xlsx.append(root)

    # 输出结果
    print("\n[ 缺失“完成报价” .docx 文档的文件夹 ]")
    if missing_docx:
        for path in missing_docx:
            print(f"- {path}")
    else:
        print("未发现缺失的情况。")

    print("\n" + "=" * 50)

    print("\n[ 缺失“完成报价” .xlsx 文档的文件夹 ]")
    if missing_xlsx:
        for path in missing_xlsx:
            print(f"- {path}")
    else:
        print("未发现缺失的情况。")
    
    print("\n" + "-" * 50)
    print(f"扫描完成。共检查了包含“上传版”的文件夹。")

if __name__ == "__main__":
    target_path = "/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/国网安徽电力第三片区2025年第十一次服务区域联合授权竞争性谈判采购_采购文件包"
    scan_missing_quote_files(target_path)
