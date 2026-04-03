import os
import threading
import json
import re
import base64
import time
import shutil
from openai import OpenAI
#本脚本作用：批量识别报价图片中的文字，获取所需文字并修改图片名称
# 配置
FOLDER_PATH = r"c:\Users\Pang Yu\Desktop\报价图片"
OUTPUT_FILE = r"c:\Users\Pang Yu\Desktop\报价图片\ocr_results.json"
MAX_THREADS = 5

# API配置（从参考脚本中复制）
API_KEY = "sk-0376a300c8444a3b84a8d725a9a39d10"
BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
MODEL = "qwen-vl-plus"

# 初始化OpenAI客户端
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

# 线程锁
lock = threading.Lock()
results = []

# 将图片转换为base64
def image_to_base64(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode("utf-8")

# 调用OpenAI API进行OCR
def ocr_image(image_path):
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            # 获取图片的base64编码
            base64_image = image_to_base64(image_path)
            
            # 调用API进行OCR
            response = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": "请识别图片中的所有文字，直接返回识别结果，不要添加任何解释。"
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{base64_image}"
                                }
                            }
                        ]
                    }
                ],
                max_tokens=2048
            )
            
            # 提取识别结果
            ocr_text = response.choices[0].message.content
            
            # 提取项目名称
            project_name = extract_project_name(ocr_text)
            
            # 保存结果
            result = {
                "image_path": image_path,
                "project_name": project_name,
                "ocr_text": ocr_text,
                "status": "success"
            }
            
            # 复制图片到指定文件夹并重命名
            try:
                # 目标文件夹路径
                img_output_dir = r"c:\Users\Pang Yu\Desktop\报价图片\ocr_img"
                # 确保文件夹存在
                os.makedirs(img_output_dir, exist_ok=True)
                
                # 获取图片扩展名
                img_ext = os.path.splitext(image_path)[1]
                
                # 使用项目名称作为新文件名
                if project_name:
                    # 清理文件名，移除非法字符
                    new_img_name = re.sub(r'[<>:"|?*\\/\r\n]', '', project_name)
                    # 替换空格为下划线
                    new_img_name = new_img_name.replace(' ', '_')
                    # 添加扩展名
                    new_img_name = new_img_name + img_ext
                else:
                    # 如果没有识别到项目名称，使用原文件名
                    new_img_name = os.path.basename(image_path)
                
                # 构建目标图片路径
                target_img_path = os.path.join(img_output_dir, new_img_name)
                
                # 复制图片
                shutil.copy2(image_path, target_img_path)
                print(f"图片已复制并重命名: {image_path} -> {target_img_path}")
            except Exception as e:
                print(f"复制图片失败: {image_path}, 错误: {e}")
            
            with lock:
                results.append(result)
            
            print(f"成功识别: {image_path}, 项目名称: {project_name}")
            return
            
        except Exception as e:
            # 检查是否为连接错误
            if "Connection error" in str(e) or "connection error" in str(e):
                retry_count += 1
                if retry_count < max_retries:
                    print(f"连接错误，{retry_count}秒后重试 ({retry_count}/{max_retries}): {image_path}")
                    time.sleep(3)
                else:
                    result = {
                        "image_path": image_path,
                        "error": str(e),
                        "status": "failed"
                    }
                    with lock:
                        results.append(result)
                    print(f"识别失败（重试{max_retries}次后）: {image_path}, 错误: {e}")
            else:
                # 其他错误直接失败
                result = {
                    "image_path": image_path,
                    "error": str(e),
                    "status": "failed"
                }
                with lock:
                    results.append(result)
                print(f"识别失败: {image_path}, 错误: {e}")
                return

# 提取项目名称
def extract_project_name(ocr_text):
    try:
        # 优化文本：移除换行符，替换英文括号为中文括号
        optimized_text = ocr_text.replace('\n', ' ').replace('\r', '')
        optimized_text = optimized_text.replace('(', '（').replace(')', '）')
        
        # 使用多种模式尝试提取项目名称
        patterns = [
            # 模式1：匹配"项目名称"或"项目名"后面的内容
            r'项目名称[：:]?\s*(.+?)[\s，。；；]+',
            r'项目名[：:]?\s*(.+?)[\s，。；；]+',
            # 模式2：匹配"全权办理"和"（项目名称）"之间的内容
            r'全权办理\s*(.+?)\s*（项目名称）',
            # 模式3：匹配"采购人"前面的内容（可能是项目名称）
            r'(.+?)\s*采购人[：:]',
            # 模式4：匹配"招标"相关的项目名称
            r'(招标项目|采购项目)名称[：:]?\s*(.+?)[\s，。；；]+'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, optimized_text, re.DOTALL | re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        # 如果没有匹配到，返回空字符串
        return ""
    except Exception as e:
        print(f"提取项目名称失败: {e}")
        return ""

# 线程工作函数
def worker(image_queue):
    while True:
        try:
            image_path = image_queue.pop()
            ocr_image(image_path)
        except IndexError:
            break

# 保存结果
def save_results(results):
    # 确保输出目录存在
    output_dir = os.path.dirname(OUTPUT_FILE)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # 保存为JSON文件
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    # 同时保存为纯文本文件，方便查看
    plain_text_file = os.path.join(output_dir, "ocr_results.txt")
    with open(plain_text_file, "w", encoding="utf-8") as f:
        for result in results:
            f.write(f"图片路径: {result['image_path']}\n")
            if result['status'] == 'success':
                f.write(f"项目名称: {result['project_name']}\n")
                f.write(f"识别文本: {result['ocr_text'][:500]}...\n")  # 只显示前500字符
            else:
                f.write(f"错误信息: {result['error']}\n")
            f.write("=" * 80 + "\n\n")
    
    print(f"结果已保存到: {OUTPUT_FILE}")
    print(f"纯文本结果已保存到: {plain_text_file}")

# 主函数
def main():
    # 获取所有支持格式的图片
    image_files = []
    for f in os.listdir(FOLDER_PATH):
        if f.lower().endswith((".jpg", ".jpeg", ".png")):
            image_files.append(os.path.join(FOLDER_PATH, f))
    
    print(f"发现 {len(image_files)} 张图片")
    
    if not image_files:
        print("未找到图片文件")
        return
    
    # 创建线程池
    threads = []
    image_queue = image_files.copy()
    
    # 启动线程
    for _ in range(min(MAX_THREADS, len(image_files))):
        thread = threading.Thread(target=worker, args=(image_queue,))
        threads.append(thread)
        thread.start()
    
    # 等待所有线程完成
    for thread in threads:
        thread.join()
    
    print(f"OCR识别完成")
    print(f"成功: {sum(1 for r in results if r['status'] == 'success')}")
    print(f"失败: {sum(1 for r in results if r['status'] == 'failed')}")
    
    # 保存结果
    save_results(results)

if __name__ == "__main__":
    main()