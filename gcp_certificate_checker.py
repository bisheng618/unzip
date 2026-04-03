import time
from openai import OpenAI
from openpyxl import load_workbook

# 配置
EXCEL_FILE = "/Users/bisheng/Downloads/Result_118.xlsx"
API_KEY = "sk-0376a300c8444a3b84a8d725a9a39d10"
BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
MODEL = "qwen-vl-plus"

# 初始化OpenAI客户端
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)


def check_gcp_certificate(image_url):
    """调用大模型判断是否为GCP证书"""
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            # 调用API进行判断,直接使用图片URL
            response = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": "请判断这张图片是否为GCP证书。请只回答'是'或'否',不要添加任何其他解释。"
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": image_url
                                }
                            }
                        ]
                    }
                ],
                max_tokens=512
            )
            
            # 提取判断结果
            result = response.choices[0].message.content.strip()
            
            # 标准化结果
            if "是" in result or "yes" in result.lower() or "gcp" in result.lower():
                return "是"
            else:
                return "否"
            
        except Exception as e:
            # 检查是否为连接错误
            if "Connection error" in str(e) or "connection error" in str(e):
                retry_count += 1
                if retry_count < max_retries:
                    print(f"连接错误,{retry_count}秒后重试 ({retry_count}/{max_retries}): {image_url}")
                    time.sleep(3)
                else:
                    print(f"判断失败(重试{max_retries}次后): {image_url}, 错误: {e}")
                    return "错误"
            else:
                print(f"判断失败: {image_url}, 错误: {e}")
                return "错误"
    
    return "错误"


def process_excel():
    """处理Excel文件"""
    # 加载工作簿
    print(f"正在加载Excel文件: {EXCEL_FILE}")
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    # 查找"path"列的索引
    path_col_idx = None
    header_row = 1
    
    for col_idx, cell in enumerate(ws[header_row], start=1):
        if cell.value and "path" in str(cell.value).lower():
            path_col_idx = col_idx
            break
    
    if path_col_idx is None:
        print("错误: 未找到'path'列")
        return
    
    print(f"找到'path'列,位于第{path_col_idx}列")
    
    # 结果列索引(path列的下一列)
    result_col_idx = path_col_idx + 1
    
    # 在表头添加结果列名称
    ws.cell(row=header_row, column=result_col_idx, value="是否GCP证书")
    
    # 遍历每一行(从第2行开始,跳过表头)
    total_rows = ws.max_row
    print(f"共有{total_rows - 1}行数据需要处理")
    
    for row_idx in range(2, total_rows + 1):
        path_cell = ws.cell(row=row_idx, column=path_col_idx)
        image_url = path_cell.value
        
        if not image_url:
            print(f"第{row_idx}行: path为空,跳过")
            ws.cell(row=row_idx, column=result_col_idx, value="无URL")
            continue
        
        print(f"第{row_idx}行/{total_rows}: 处理图片 {image_url}")
        
        # 直接使用URL判断是否为GCP证书
        result = check_gcp_certificate(image_url)
        print(f"第{row_idx}行: 判断结果 = {result}")
        
        # 写入结果
        ws.cell(row=row_idx, column=result_col_idx, value=result)
        
        # 每处理10行保存一次,防止数据丢失
        if row_idx % 10 == 0:
            wb.save(EXCEL_FILE)
            print(f"已保存进度: {row_idx}/{total_rows}")
    
    # 最终保存
    wb.save(EXCEL_FILE)
    print(f"\n处理完成!结果已保存到: {EXCEL_FILE}")


def main():
    """主函数"""
    print("=" * 60)
    print("GCP证书检测程序")
    print("=" * 60)
    
    try:
        process_excel()
        print("\n程序执行完毕!")
    except FileNotFoundError:
        print(f"错误: Excel文件不存在: {EXCEL_FILE}")
    except Exception as e:
        print(f"程序执行出错: {e}")


if __name__ == "__main__":
    main()
