#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import glob
import time
import subprocess
import fitz  # PyMuPDF
from PIL import Image, ImageChops
import random

def convert_docx_to_pdf(docx_path, pdf_path):
    """Use LibreOffice to convert DOCX to PDF"""
    output_dir = os.path.dirname(pdf_path)
    # Ensure absolute paths
    docx_path = os.path.abspath(docx_path)
    output_dir = os.path.abspath(output_dir)
    
    cmd = [
        '/opt/homebrew/bin/soffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', output_dir,
        docx_path
    ]
    
    try:
        subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        # LibreOffice saves with the same basename
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        generated_pdf = os.path.join(output_dir, f"{base_name}.pdf")
        
        # Wait for file
        max_wait = 10
        wait_time = 0
        while not os.path.exists(generated_pdf) and wait_time < max_wait:
            time.sleep(0.5)
            wait_time += 0.5
            
        if not os.path.exists(generated_pdf):
            print(f"Error: PDF not generated for {docx_path}")
            return False
            
        # Rename if necessary (though usually we just want the pdf at the specific path)
        if generated_pdf != pdf_path:
            os.rename(generated_pdf, pdf_path)
            
        return True
    except Exception as e:
        print(f"Conversion error: {e}")
        return False

def process_document(docx_path, stamp_path):
    print(f"Processing: {os.path.basename(docx_path)}")
    
    # 1. Convert to temporary PDF
    pdf_path = docx_path.replace(".docx", "_temp.pdf")
    if not convert_docx_to_pdf(docx_path, pdf_path):
        return

    try:
        doc = fitz.open(pdf_path)
        
        # Pages to process: 5 and 7 (indices 4 and 6)
        target_pages = [4]
        
        for page_idx in target_pages:
            if page_idx >= len(doc):
                print(f"  Skipping page {page_idx+1} (Document has only {len(doc)} pages)")
                continue
                
            page = doc[page_idx]
            
            # 2. Find "盖章" text coordinates
            # search_for returns list of Rect objects
            text_instances = page.search_for("盖章")
            
            # 3. Render page to Image (High DPI for quality)
            zoom = 300 / 72  # 72 dpi -> 300 dpi
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # 4. Stamp the image
            if text_instances:
                stamp_img = Image.open(stamp_path).convert("RGBA")
                
                # Resize stamp (e.g., 4cm / 1.57 inch width)
                # Let's use 1.5 inch width as standard for stamps
                target_width_inch = 1.5
                target_width_px = int(target_width_inch * 300)
                aspect_ratio = stamp_img.height / stamp_img.width
                target_height_px = int(target_width_px * aspect_ratio)
                
                stamp_img = stamp_img.resize((target_width_px, target_height_px), Image.Resampling.LANCZOS)
                
                print(f"  Found {len(text_instances)} '盖章' on page {page_idx+1}")
                
                for rect in text_instances:
                    # Rect is (x0, y0, x1, y1) in points (72 dpi)
                    # We need to map center of rect to image pixels
                    
                    center_x_pt = (rect.x0 + rect.x1) / 2
                    center_y_pt = (rect.y0 + rect.y1) / 2
                    
                    center_x_px = int(center_x_pt * zoom)
                    center_y_px = int(center_y_pt * zoom)
                    
                    # Realism 1: Random Rotation (-5 to 5 degrees)
                    angle = random.uniform(-5, 5)
                    stamp_rotated = stamp_img.rotate(angle, resample=Image.Resampling.BICUBIC, expand=True)
                    
                    # Calculate top-left for paste (adjusting for rotated size)
                    paste_x = center_x_px - (stamp_rotated.width // 2)
                    paste_y = center_y_px - (stamp_rotated.height // 2)
                    
                    # Realism 2: Multiply Blending (Simulate ink absorption)
                    # Create a white layer with the stamp
                    stamp_layer = Image.new('RGB', img.size, (255, 255, 255))
                    # Paste stamp onto white layer using its alpha channel as mask
                    stamp_layer.paste(stamp_rotated, (paste_x, paste_y), stamp_rotated)
                    
                    # Multiply original image with stamp layer
                    # (White * Color = Color, Black * Color = Black)
                    img = ImageChops.multiply(img, stamp_layer)
            else:
                print(f"  No '盖章' found on page {page_idx+1}")

            # 5. Save Image
            # Format: filename_page5.png
            base_name = os.path.splitext(os.path.basename(docx_path))[0]
            output_filename = f"{base_name}_page{page_idx+1}.png"
            output_path = os.path.join(os.path.dirname(docx_path), output_filename)
            
            img.save(output_path)
            print(f"  Saved: {output_filename}")

        doc.close()
        
    except Exception as e:
        print(f"Error processing {docx_path}: {e}")
    finally:
        # Cleanup temp PDF
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

def main():
    source_dir = "/Users/bisheng/Downloads/小玲模板/第一片区/"
    stamp_path = "/Users/bisheng/Downloads/yide.png"
    
    if not os.path.exists(source_dir):
        print(f"Source directory not found: {source_dir}")
        return
        
    if not os.path.exists(stamp_path):
        print(f"Stamp image not found: {stamp_path}")
        return
        
    docx_files = glob.glob(os.path.join(source_dir, "*.docx"))
    # Filter out temp files
    docx_files = [f for f in docx_files if not os.path.basename(f).startswith("~$")]
    
    print(f"Found {len(docx_files)} docx files.")
    
    for docx_file in docx_files:
        process_document(docx_file, stamp_path)
        
    print("Done.")

if __name__ == "__main__":
    main()
