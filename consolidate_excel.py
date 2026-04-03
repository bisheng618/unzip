import pandas as pd
import os
from openpyxl import load_workbook

def consolidate_excel_sheets(file_path, output_sheet_name='汇总'):
    """
    Consolidates all sheets in an Excel file into a single summary sheet.
    
    Args:
        file_path (str): Path to the Excel file.
        output_sheet_name (str): Name of the sheet to store consolidated data.
    """
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    print(f"Processing file: {file_path}")

    try:
        # Read all sheets from the Excel file
        xls = pd.ExcelFile(file_path)
        all_sheets = xls.sheet_names
        
        # Filter out the output sheet if it already exists to avoid self-inclusion
        sheets_to_process = [s for s in all_sheets if s != output_sheet_name]
        
        if not sheets_to_process:
            print("No sheets to consolidate.")
            return

        print(f"Found {len(sheets_to_process)} sheets to consolidate.")
        
        # List to hold dataframes
        df_list = []
        
        for sheet_name in sheets_to_process:
            print(f"Reading sheet: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
            # Add a column to identify the source sheet (optional but helpful)
            df['来源表格'] = sheet_name 
            df_list.append(df)
            
        if not df_list:
            print("No data found in sheets.")
            return

        # Concatenate all dataframes
        consolidated_df = pd.concat(df_list, ignore_index=True)
        
        print(f"Consolidated {len(consolidated_df)} rows.")

        # Write to the same Excel file
        # use openpyxl engine to append mode requires keeping existing sheets?
        # Actually, pd.ExcelWriter with mode='a' and if_sheet_exists='replace' is best for modern pandas
        
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            consolidated_df.to_excel(writer, sheet_name=output_sheet_name, index=False)
            
        print(f"Successfully consolidated data into sheet '{output_sheet_name}' in {file_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    target_file = "/Users/bisheng/Downloads/采购公告 (3)/货物清单_国网安徽电力第一片区2025年第十一次服务区域联合授权竞争性谈判采购20251216_214339_344.xlsx"
    consolidate_excel_sheets(target_file)
