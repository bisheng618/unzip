
import os
from PIL import Image

def resize_and_pad_image(image, target_width, target_height):
    """
    Resizes an image to fit within target dimensions while maintaining aspect ratio,
    and then pads it with white background to reach the target_width x target_height.
    """
    original_width, original_height = image.size
    
    # Calculate aspect ratios
    target_aspect = target_width / target_height
    image_aspect = original_width / original_height

    if image_aspect > target_aspect:
        # Image is wider than target aspect ratio, scale by width
        new_width = target_width
        new_height = int(target_width / image_aspect)
    else:
        # Image is taller or same aspect ratio, scale by height
        new_height = target_height
        new_width = int(target_height * image_aspect)

    # Resize the image
    resized_image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)

    # Create a new blank white image (canvas)
    padded_image = Image.new("RGB", (target_width, target_height), (255, 255, 255)) # White background

    # Calculate paste position to center the resized image
    paste_x = (target_width - new_width) // 2
    paste_y = (target_height - new_height) // 2

    # Paste the resized image onto the center of the canvas
    padded_image.paste(resized_image, (paste_x, paste_y))
    
    return padded_image

def create_multipage_pdfs():
    """
    Reads all images from a source directory, combines each with a static image,
    and saves them as two-page PDF files, ensuring uniform page sizes.
    """
    # --- 配置路径 ---
    # 1. 包含动态图片的文件夹
    source_image_dir = "/Users/bisheng/Downloads/证件/法定代表人授权委托书/授权委托书"
    
    # 2. 固定的第二页图片路径
    static_image_path = "/Users/bisheng/Downloads/证件/sfz.jpg"
    
    # 3. PDF输出文件夹
    output_pdf_dir = "/Users/bisheng/Downloads/证件/法定代表人授权委托书/generated_pdfs"
    
    # --- 检查路径是否存在 ---
    if not os.path.isdir(source_image_dir):
        print(f"错误: 来源图片文件夹不存在: {source_image_dir}")
        return

    if not os.path.isfile(static_image_path):
        print(f"错误: 静态图片文件不存在: {static_image_path}")
        return
        
    # 如果输出文件夹不存在,则创建它
    if not os.path.exists(output_pdf_dir):
        os.makedirs(output_pdf_dir)
        print(f"已创建输出文件夹: {output_pdf_dir}")

    # --- 打开静态图片并确定目标尺寸 ---
    try:
        img_static = Image.open(static_image_path)
        # 转换为RGB以确保兼容性
        img_static_rgb_original = img_static.convert("RGB")
        target_width, target_height = img_static_rgb_original.size
        print(f"以静态图片 '{static_image_path}' 的尺寸 ({target_width}x{target_height}) 作为PDF页面目标尺寸。")
    except Exception as e:
        print(f"无法打开静态图片 '{static_image_path}' 或获取其尺寸: {e}")
        # 提供一个默认的A4比例尺寸 (例如: 150 DPI 的 A4 页面)
        target_width, target_height = 1240, 1754 
        print(f"使用默认尺寸 {target_width}x{target_height} 作为PDF页面目标尺寸。")
    
    # 调整静态图片到目标尺寸
    img_static_processed = resize_and_pad_image(img_static_rgb_original, target_width, target_height)

    # --- 遍历源文件夹中的图片并生成PDF ---
    print(f"开始处理文件夹 '{source_image_dir}' 中的图片...")
    
    found_images = False
    for filename in os.listdir(source_image_dir):
        # 检查是否是常见的图片格式
        if filename.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".gif")):
            found_images = True
            source_image_path = os.path.join(source_image_dir, filename)
            
            try:
                # 打开源图片
                img_source = Image.open(source_image_path)
                # 转换为RGB
                img_source_rgb_original = img_source.convert("RGB")
                
                # 调整源图片到目标尺寸
                img_source_processed = resize_and_pad_image(img_source_rgb_original, target_width, target_height)
                
                # 定义PDF输出路径
                pdf_filename = os.path.splitext(filename)[0] + ".pdf"
                output_pdf_path = os.path.join(output_pdf_dir, pdf_filename)
                
                # 准备图片列表用于合并 (第一页是源图片, 第二页是静态图片)
                images_to_save = [img_static_processed] # 第二页
                
                # 保存为多页PDF
                img_source_processed.save(
                    output_pdf_path,
                    "PDF",
                    resolution=100.0, # 可以根据需要调整分辨率
                    save_all=True,
                    append_images=images_to_save
                )
                
                print(f"成功创建PDF: {output_pdf_path}")
                
            except Exception as e:
                print(f"处理图片 '{filename}' 时出错: {e}")

    if not found_images:
        print("在源文件夹中没有找到任何图片文件。 ")
    else:
        print("\n所有图片处理完毕。")

if __name__ == "__main__":
    create_multipage_pdfs()
