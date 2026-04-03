import os
import shutil

# 源目录（已挂载）
SRC_DIR = r"/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/"

# 目标目录
DEST_DIR = "/Users/bisheng/Downloads/11281"

# TXT 输出路径
TXT_PATH = os.path.join(DEST_DIR, "files.txt")


def copy_excel_files_and_generate_txt():
    if not os.path.exists(SRC_DIR):
        print("❌ 源目录不存在，请检查路径：", SRC_DIR)
        return

    os.makedirs(DEST_DIR, exist_ok=True)

    entries = []
    copied_count = 0

    for root, dirs, files in os.walk(SRC_DIR):
        for file in files:
            # 匹配包含“完成报价”的 Excel
            if "完成报价" in file and file.lower().endswith((".xlsx", ".xls")):
                src_path = os.path.join(root, file)
                dest_path = os.path.join(DEST_DIR, file)

                print(f"复制：{src_path} -> {dest_path}")

                shutil.copy2(src_path, dest_path)
                copied_count += 1

                # 构造 "目录|文件名" 格式
                entries.append(f"{root}|{file}")

    # 写入 TXT
    with open(TXT_PATH, "w", encoding="utf-8") as f:
        for line in entries:
            f.write(line + "\n")

    print(f"\n🎉 完成！共复制 {copied_count} 个 Excel 文件。")
    print(f"📄 TXT 文件已生成：{TXT_PATH}")


if __name__ == "__main__":
    copy_excel_files_and_generate_txt()