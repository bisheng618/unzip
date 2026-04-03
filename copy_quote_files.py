import os
import shutil
from pathlib import Path

def copy_quote_files(source_root, dest_root):
    """
    Recursively copies files containing "完成报价" from source_root to dest_root,
    preserving the directory structure relative to source_root.
    """
    source_path = Path(source_root)
    dest_path = Path(dest_root)

    if not source_path.exists():
        print(f"Error: Source directory does not exist: {source_path}")
        return

    print(f"Scanning source directory: {source_path}")
    print(f"Target destination: {dest_path}")

    copied_count = 0
    
    import re
    
    def transform_path(relative_path):
        """
        Transforms paths to map directory names like '包4_...' to '包04'.
        """
        parts = relative_path.parts
        new_parts = []
        for part in parts:
            # Check if the part starts with '包' followed by digits and an underscore
            match = re.match(r'^包(\d+)_', part)
            if match:
                # Extract the number and format it to 2 digits (e.g., 4 -> 04)
                pkg_num = int(match.group(1))
                new_part = f"包{pkg_num:02d}"
                new_parts.append(new_part)
            else:
                new_parts.append(part)
        return Path(*new_parts)

    for root, dirs, files in os.walk(source_path):
        for file in files:
            if "完成报价" in file:
                # Construct full source file path
                src_file_path = Path(root) / file
                
                # Calculate relative path from source root
                relative_path = src_file_path.relative_to(source_path)
                
                # Transform path (map '包4_...' to '包04')
                transformed_relative_path = transform_path(relative_path)
                
                # Construct destination file path
                dest_file_path = dest_path / transformed_relative_path
                
                # Create destination directory if it doesn't exist
                dest_file_path.parent.mkdir(parents=True, exist_ok=True)
                
                try:
                    print(f"Copying: {transformed_relative_path}")
                    shutil.copy2(src_file_path, dest_file_path)
                    copied_count += 1
                except Exception as e:
                    print(f"Failed to copy {src_file_path}: {e}")

    print(f"\nOperation complete. Copied {copied_count} files.")

if __name__ == "__main__":
    SOURCE_DIR = "/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/国网安徽电力第三片区2025年第十次服务区域联合授权竞争性谈判采购-2025-12-02-AM6：00-ECP/国网安徽电力第三片区2025年第十次服务区域联合授权竞争性谈判采购_采购文件包"
    DEST_DIR = "/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/bisheng测试/国网安徽电力第三片区2025年第十次服务区域联合授权竞争性谈判采购_采购文件包"
    
    copy_quote_files(SOURCE_DIR, DEST_DIR)
