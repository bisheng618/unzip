import os
from pathlib import Path
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def adjust_side_by_side_images(doc_path, output_path, left_margin_cm=2.54, right_margin_cm=2.54, safety_margin_cm=0.3):
    """
    自动调整Word文档中的图片,确保两张图片能并排显示
    只处理特定章节内的图片
    """
    
    doc = Document(doc_path)
    
    # 获取页面宽度(A4纸张宽度为21cm)
    page_width_cm = 21.0
    
    # 计算可用宽度,预留安全余量
    available_width_cm = page_width_cm - left_margin_cm - right_margin_cm - safety_margin_cm
    
    # 计算每张图片的最大宽度
    max_image_width_cm = (available_width_cm - 0.5) / 2
    
    print(f"页面宽度: {page_width_cm} cm")
    print(f"可用宽度: {available_width_cm:.2f} cm")
    print(f"每张图片最大宽度: {max_image_width_cm:.2f} cm\n")
    
    # 定义需要处理的章节标题
    target_sections = ["拟委任的主要人员汇总表", "拟配备的本项目服务人员"]
    
    # 查找目标章节的范围
    section_ranges = []
    current_section_start = None
    current_section_name = None
    
    for para_index, para in enumerate(doc.paragraphs):
        # 检查是否是章节标题
        text = para.text.strip()
        for section_name in target_sections:
            if section_name in text:
                # 如果已有未闭合的章节,先保存
                if current_section_start is not None:
                    section_ranges.append({
                        'name': current_section_name,
                        'start': current_section_start,
                        'end': para_index - 1
                    })
                current_section_start = para_index
                current_section_name = section_name
                print(f"找到章节: {section_name} (行 {para_index + 1})")
                break
    
    # 保存最后一个章节
    if current_section_start is not None:
        section_ranges.append({
            'name': current_section_name,
            'start': current_section_start,
            'end': len(doc.paragraphs) - 1
        })
    
    print(f"共找到 {len(section_ranges)} 个目标章节\n")
    
    # 用于标记需要处理的段落
    paragraphs_to_process = []
    
    # 遍历所有段落
    for para_index, para in enumerate(doc.paragraphs):
        # 检查该段落是否在目标章节范围内
        in_target_section = False
        for section in section_ranges:
            if section['start'] <= para_index <= section['end']:
                in_target_section = True
                break
        
        # 只处理在目标章节内的段落
        if not in_target_section:
            continue
        
        images_info = []
        
        for run in para.runs:
            # 查找drawing元素
            for drawing in run._element.findall(qn('w:drawing')):
                # 查找inline元素
                inline = drawing.find(qn('wp:inline'))
                if inline is not None:
                    # 查找extent元素
                    extent = inline.find(qn('wp:extent'))
                    if extent is not None:
                        images_info.append({
                            'run': run,
                            'drawing': drawing,
                            'inline': inline,
                            'extent': extent
                        })
        
        if len(images_info) >= 2:
            # 检查图片是否会导致换行(即超过可用宽度)
            total_width_cm = 0
            for img in images_info[:2]:
                extent = img['extent']
                width_emu = extent.get('cx')
                if width_emu:
                    total_width_cm += int(width_emu) / 360000
            
            # 如果两张图片总宽度超过可用宽度,才需要处理
            if total_width_cm > available_width_cm:
                paragraphs_to_process.append({
                    'para': para,
                    'para_index': para_index,
                    'images': images_info
                })
    
    # 处理找到的段落
    processed_count = 0
    for item in reversed(paragraphs_to_process):
        para = item['para']
        para_index = item['para_index']
        images = item['images'][:2]  # 只取前两张
        
        print(f"处理段落 {para_index + 1}, 包含 {len(item['images'])} 张图片")
        
        try:
            # 调整图片尺寸
            adjusted_images = []
            for i, img_info in enumerate(images):
                extent = img_info['extent']
                
                # 获取当前尺寸 (EMU单位)
                current_width_emu = extent.get('cx')
                current_height_emu = extent.get('cy')
                
                if current_width_emu is None or current_height_emu is None:
                    print(f"  警告: 图片 {i+1} 无法获取尺寸信息,跳过")
                    continue
                
                current_width_emu = int(current_width_emu)
                current_height_emu = int(current_height_emu)
                
                # 转换为厘米
                current_width_cm = current_width_emu / 360000
                current_height_cm = current_height_emu / 360000
                
                print(f"  图片 {i+1} 原始: {current_width_cm:.2f} x {current_height_cm:.2f} cm", end="")
                
                # 计算新尺寸
                if current_width_cm > max_image_width_cm:
                    scale_factor = max_image_width_cm / current_width_cm
                    new_width_emu = int(current_width_emu * scale_factor)
                    new_height_emu = int(current_height_emu * scale_factor)
                    
                    new_width_cm = new_width_emu / 360000
                    new_height_cm = new_height_emu / 360000
                    
                    # 更新extent
                    extent.set('cx', str(new_width_emu))
                    extent.set('cy', str(new_height_emu))
                    
                    # 更新ext元素
                    ext = img_info['inline'].find(qn('a:graphic')).find(qn('a:graphicData')).find(qn('pic:pic'))
                    if ext is not None:
                        xfrm = ext.find(qn('pic:spPr')).find(qn('a:xfrm'))
                        if xfrm is not None:
                            ext_elem = xfrm.find(qn('a:ext'))
                            if ext_elem is not None:
                                ext_elem.set('cx', str(new_width_emu))
                                ext_elem.set('cy', str(new_height_emu))
                    
                    print(f" -> 调整为: {new_width_cm:.2f} x {new_height_cm:.2f} cm ✓")
                else:
                    new_width_emu = current_width_emu
                    new_height_emu = current_height_emu
                    print(f" -> 无需调整 ✓")
                
                adjusted_images.append(img_info)
            
            if len(adjusted_images) < 2:
                print(f"  跳过此段落:图片数量不足\n")
                continue
            
            # 创建表格
            table = doc.add_table(rows=1, cols=2)
            table.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 设置表格样式
            tbl = table._element
            tblPr = tbl.find(qn('w:tblPr'))
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                tbl.insert(0, tblPr)
            
            # 移除边框
            tblBorders = OxmlElement('w:tblBorders')
            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'none')
                border.set(qn('w:sz'), '0')
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), 'auto')
                tblBorders.append(border)
            tblPr.append(tblBorders)
            
            # 设置表格宽度和单元格宽度
            tblW = OxmlElement('w:tblW')
            tblW.set(qn('w:w'), '5000')  # 自动宽度
            tblW.set(qn('w:type'), 'auto')
            tblPr.append(tblW)
            
            # 将图片移到表格单元格
            for i, img_info in enumerate(adjusted_images):
                cell = table.rows[0].cells[i]
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # 移动drawing到单元格
                cell.paragraphs[0]._element.append(img_info['drawing'])
            
            # 在原段落位置插入表格
            para._element.addnext(table._element)
            
            # 删除原段落
            p = para._element
            p.getparent().remove(p)
            
            processed_count += 1
            print()
            
        except Exception as e:
            print(f"  错误: {str(e)}\n")
            continue
    
    # 设置页边距
    section = doc.sections[0]
    section.left_margin = Cm(left_margin_cm)
    section.right_margin = Cm(right_margin_cm)
    
    # 保存
    doc.save(output_path)
    return processed_count


def process_directory(directory_path):
    """
    扫描目录下所有包含"技术方案"的docx文件进行处理
    """
    
    directory = Path(directory_path)
    
    if not directory.exists():
        print(f"❌ 目录不存在: {directory_path}")
        return
    
    # 查找所有包含"技术方案"的docx文件(包括子文件夹)
    docx_files = list(directory.rglob("*技术方案*.docx"))
    
    if not docx_files:
        print(f"⚠️  未找到包含'技术方案'的docx文件")
        return
    
    print(f"✓ 找到 {len(docx_files)} 个文件\n")
    print("=" * 80)
    
    for file_path in docx_files:
        print(f"\n📄 处理文件: {file_path.name}")
        print("-" * 80)
        
        # 生成输出文件名(保存到原文件所在目录)
        output_filename = f"修复后_{file_path.name}"
        output_path = file_path.parent / output_filename
        
        try:
            # 处理文档
            processed_count = adjust_side_by_side_images(
                doc_path=str(file_path),
                output_path=str(output_path),
                left_margin_cm=2.54,
                right_margin_cm=2.54,
                safety_margin_cm=0.3
            )
            
            print(f"✓ 文件已保存: {output_filename}")
            print(f"✓ 共处理了 {processed_count} 个段落\n")
            
        except Exception as e:
            print(f"❌ 处理失败: {str(e)}\n")
    
    print("=" * 80)
    print("✓ 所有文件处理完成!")


if __name__ == "__main__":
    # 扫描的目录路径
    target_directory = "/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/国网安徽电力第二片区2025年第十次服务区域联合授权竞争性谈判采购-2025-12-02-AM6：00-ECP/"
    
    process_directory(target_directory)