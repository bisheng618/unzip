#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word文档拆分工具 - 定制版 (zhongcheng)
直接操作Word文档的XML结构，完全保留原始格式
"""

import os
import subprocess
import shutil
import re
from docx import Document
import tempfile

class DocxSplitterV2:
    def __init__(self, source_file, output_dir):
        self.source_file = source_file
        self.output_dir = output_dir
        self.doc = Document(source_file)
        
    def get_paragraph_style_level(self, paragraph):
        """获取段落的标题级别"""
        style_name = paragraph.style.name
        if 'Heading' in style_name or '标题' in style_name:
            match = re.search(r'(\d+)', style_name)
            if match:
                return int(match.group(1))
        return None
    
    def find_end_of_section(self, start_idx):
        """找到章节的结束位置（下一个同级或更高级别标题的开始）"""
        start_para = self.doc.paragraphs[start_idx]
        start_level = self.get_paragraph_style_level(start_para)
        
        # 如果起始段落没有级别（不是标题），则一直到文档末尾或遇到任何标题
        # 但通常我们是针对标题调用的
        
        for i in range(start_idx + 1, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            level = self.get_paragraph_style_level(para)
            
            if level is not None:
                # 如果start_level存在，我们找<=start_level的
                if start_level is not None:
                    if level <= start_level:
                        return i
                # 如果start_level不存在（正文开始），遇到任何标题可能都算结束？
                # 这里假设我们总是从标题开始找
        
        return len(self.doc.paragraphs)

    def find_subsections_by_parent(self, parent_para_idx):
        """找到某个段落下的所有直接子章节"""
        sections = []
        
        # 获取父段落的级别
        parent_para = self.doc.paragraphs[parent_para_idx]
        parent_level = self.get_paragraph_style_level(parent_para)
        
        if parent_level is None:
            # 如果父段落不是标题，无法确定子章节级别，尝试自动探测
            # 这里简单处理：假设下一级是遇到的第一个标题级别
            target_level = None
        else:
            target_level = None # 我们不知道下一级是几，可能是parent_level + 1，也可能是 + 2
        
        current_section = None
        
        for i in range(parent_para_idx + 1, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            level = self.get_paragraph_style_level(para)
            
            if level is not None:
                # 如果遇到同级或更高级别的标题，说明父章节结束了
                if parent_level is not None and level <= parent_level:
                    if current_section:
                        current_section['end_idx'] = i
                        sections.append(current_section)
                    break
                
                # 确定目标子级别
                if target_level is None:
                    target_level = level
                
                # 如果是目标子级别，开始新section
                if level == target_level:
                    # 保存上一个section
                    if current_section:
                        current_section['end_idx'] = i
                        sections.append(current_section)
                    
                    # 开始新section
                    current_section = {
                        'title': para.text.strip(),
                        'start_idx': i,
                        'end_idx': None
                    }
                # 如果是更深级别的标题，属于当前section，不用处理
                elif level < target_level:
                    # 这种情况理论上在上面 "level <= parent_level" 已经处理了（如果target > parent）
                    # 但如果结构混乱，比如 parent(1) -> child(3) -> child(2)，这里会有问题
                    # 暂时假设文档结构是规范的
                    pass
        
        # 添加最后一个section
        if current_section and not current_section['end_idx']:
            current_section['end_idx'] = len(self.doc.paragraphs)
            sections.append(current_section)
        
        return sections
    
    def extract_docx_by_paragraph_range(self, start_para_idx, end_para_idx, output_path):
        """通过复制整个docx文件并删除不需要的段落来提取内容，完全保留格式"""
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            # 复制原始文档到临时位置
            temp_docx = os.path.join(temp_dir, 'temp.docx')
            shutil.copy2(self.source_file, temp_docx)
            
            # 打开复制的文档
            doc = Document(temp_docx)
            
            # 获取所有段落和它们的父元素
            body = doc.element.body
            
            # 收集要删除的元素
            elements_to_remove = []
            
            # Word命名空间
            w_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            
            # 先获取所有段落的样式信息
            para_styles = []
            for para in doc.paragraphs:
                para_styles.append(para.style.name.lower())
            
            # 遍历body的所有子元素
            para_count = 0
            for element in list(body):
                # 检查是否是段落
                if element.tag.endswith('p'):
                    should_remove = False
                    
                    # 检查是否在范围外
                    if para_count < start_para_idx or (end_para_idx and para_count >= end_para_idx):
                        should_remove = True
                    else:
                        # 检查段落样式是否包含'toc'（目录样式）
                        if para_count < len(para_styles):
                            style_name = para_styles[para_count]
                            if 'toc' in style_name:
                                should_remove = True
                        
                        # 检查是否包含TOC字段指令
                        if not should_remove:
                            for child in element.iter():
                                # 检查instrText元素
                                if child.tag == f'{w_namespace}instrText':
                                    if child.text and 'TOC' in child.text:
                                        should_remove = True
                                        break
                        
                        # 如果段落文本为"目录"，也删除
                        if not should_remove:
                            para_text = ''
                            for child in element.iter():
                                if child.tag == f'{w_namespace}t':
                                    if child.text:
                                        para_text += child.text
                            
                            if para_text.strip() in ['目录', 'Table of Contents', 'TOC']:
                                should_remove = True
                    
                    if should_remove:
                        elements_to_remove.append(element)
                    
                    para_count += 1
                    
                # 如果是表格，检查它是否在范围内
                elif element.tag.endswith('tbl'):
                    # 表格比较复杂，我们需要检查它前后的段落位置
                    # 为了简单起见，如果表格在段落范围内就保留
                    if para_count < start_para_idx or (end_para_idx and para_count >= end_para_idx):
                        elements_to_remove.append(element)
            
            # 删除不需要的元素
            for element in elements_to_remove:
                body.remove(element)
            
            # 保存修改后的文档
            doc.save(output_path)
    
    def convert_docx_to_pdf(self, docx_path, pdf_path):
        """使用LibreOffice将DOCX转换为PDF"""
        try:
            cmd = [
                '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', os.path.dirname(pdf_path),
                docx_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60) # 增加超时时间
            
            # LibreOffice会生成与输入文件同名的PDF
            generated_pdf = docx_path.replace('.docx', '.pdf')
            
            if os.path.exists(generated_pdf):
                if generated_pdf != pdf_path:
                    shutil.move(generated_pdf, pdf_path)
                print(f"✓ 已生成: {os.path.basename(pdf_path)}")
                return True
            else:
                print(f"✗ 转换失败: {os.path.basename(pdf_path)}")
                if result.stderr:
                    print(f"  错误信息: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            print(f"✗ 转换超时: {os.path.basename(pdf_path)}")
            return False
        except Exception as e:
            print(f"✗ 转换出错 {os.path.basename(pdf_path)}: {e}")
            return False
    
    def save_section_as_pdf(self, start_idx, end_idx, output_path):
        """保存段落范围为PDF"""
        temp_docx = output_path.replace('.pdf', '_temp.docx')
        
        # 提取内容到临时docx
        self.extract_docx_by_paragraph_range(start_idx, end_idx, temp_docx)
        
        # 转换为PDF
        success = self.convert_docx_to_pdf(temp_docx, output_path)
        
        # 删除临时文件
        if os.path.exists(temp_docx):
            os.remove(temp_docx)
        
        return success
    
    def split_related_materials(self):
        """拆分"相关证明材料按项目逐一列明"下的3个小节"""
        print("\n1. 拆分相关证明材料...")
        output_folder = os.path.join(self.output_dir, "相关证明材料")
        os.makedirs(output_folder, exist_ok=True)
        
        parent_idx = None
        for i, para in enumerate(self.doc.paragraphs):
            if "相关证明材料按项目逐一列明" in para.text:
                parent_idx = i
                break
        
        if parent_idx is None:
            print("  未找到'相关证明材料按项目逐一列明'")
            return
        
        sections = self.find_subsections_by_parent(parent_idx)
        print(f"  找到 {len(sections)} 个小节")
        
        for idx, section in enumerate(sections[:4], 1): # 只取前4个
            title = section['title']
            safe_title = re.sub(r'[\\/:*?"<>|]', '_', title)
            output_path = os.path.join(output_folder, f"{idx:02d}_{safe_title}.pdf")
            self.save_section_as_pdf(section['start_idx'], section['end_idx'], output_path)

    def split_personnel_summary(self):
        """拆分"拟委任的主要人员汇总表" """
        print("\n2. 拆分拟委任的主要人员汇总表...")
        
        start_idx = None
        for i, para in enumerate(self.doc.paragraphs):
            if "拟委任的主要人员汇总表" in para.text:
                start_idx = i
                break
        
        if start_idx is None:
            print("  未找到'拟委任的主要人员汇总表'")
            return

        end_idx = self.find_end_of_section(start_idx)
        
        output_path = os.path.join(self.output_dir, "拟委任的主要人员汇总表.pdf")
        self.save_section_as_pdf(start_idx, end_idx, output_path)

    def split_resumes(self):
        """拆分"表2主要人员简历表" """
        print("\n3. 拆分表2主要人员简历表...")
        
        start_idx = None
        for i, para in enumerate(self.doc.paragraphs):
            # 宽松匹配，因为可能有空格或编号
            if "表2" in para.text and "主要人员简历表" in para.text:
                start_idx = i
                break
        
        if start_idx is None:
            print("  未找到'表2主要人员简历表'")
            return

        end_idx = self.find_end_of_section(start_idx)
        
        output_path = os.path.join(self.output_dir, "表2主要人员简历表.pdf")
        self.save_section_as_pdf(start_idx, end_idx, output_path)

    def split_chapters_2_to_7(self):
        """拆分第二章到第七章"""
        print("\n4. 拆分第二章到第七章...")
        
        start_idx = None
        end_idx = None
        
        # 根据文档结构分析：
        # 第二章标题为 "项目概述与理解" (Level 1)
        # 第八章标题为 "验收配合与保障" (Level 1)
        start_title = "项目概述与理解"
        end_title = "验收配合与保障"
        
        for i, para in enumerate(self.doc.paragraphs):
            level = self.get_paragraph_style_level(para)
            text = para.text.strip()
            
            # 查找第二章
            if level == 1 and start_title in text:
                start_idx = i
            
            # 查找第八章（作为结束）
            if start_idx and level == 1 and end_title in text:
                end_idx = i
                break
        
        if start_idx:
            output_path = os.path.join(self.output_dir, "第二章到第七章.pdf")
            self.save_section_as_pdf(start_idx, end_idx, output_path)
        else:
            print(f"  未找到起始章节: {start_title}")

    def split_financial_status(self):
        """拆分"第一节财务状况"下的3个小节"""
        print("\n5. 拆分第一节财务状况...")
        output_folder = os.path.join(self.output_dir, "财务状况")
        os.makedirs(output_folder, exist_ok=True)
        
        parent_idx = None
        for i, para in enumerate(self.doc.paragraphs):
            # 严格匹配 Level 2 的 "财务状况"
            if self.get_paragraph_style_level(para) == 2 and "财务状况" == para.text.strip():
                parent_idx = i
                break
        
        if parent_idx is None:
            print("  未找到'财务状况' (Level 2)，尝试模糊匹配...")
            for i, para in enumerate(self.doc.paragraphs):
                if "财务状况" in para.text and self.get_paragraph_style_level(para) is not None:
                    # 检查后续段落是否包含年份，确认是正确的位置
                    is_target = False
                    for j in range(1, 5):
                        if i + j < len(self.doc.paragraphs):
                            if "2022年" in self.doc.paragraphs[i+j].text or "审计报告" in self.doc.paragraphs[i+j].text:
                                is_target = True
                                break
                    if is_target:
                        parent_idx = i
                        break
        
        if parent_idx is None:
            print("  未找到'第一节财务状况'")
            return
            
        sections = self.find_subsections_by_parent(parent_idx)
        print(f"  找到 {len(sections)} 个小节")
        
        for idx, section in enumerate(sections[:3], 1):
            title = section['title']
            safe_title = re.sub(r'[\\/:*?"<>|]', '_', title)
            output_path = os.path.join(output_folder, f"{idx:02d}_{safe_title}.pdf")
            self.save_section_as_pdf(section['start_idx'], section['end_idx'], output_path)
    
    def run(self):
        """执行所有拆分任务"""
        print("=" * 80)
        print("开始拆分文档（完全保留格式）...")
        print("=" * 80)
        
        self.split_related_materials()
        self.split_personnel_summary()
        self.split_resumes()
        self.split_chapters_2_to_7()
        self.split_financial_status()
        
        print("\n" + "=" * 80)
        print("所有拆分任务完成！")
        print("=" * 80)


def main():
    source_file = "/Users/bisheng/Downloads/筑鹏/综合服务-亳州-技术文件-包2.docx"
    output_dir = "/Users/bisheng/Downloads/筑鹏/"
    
    if not os.path.exists(source_file):
        print(f"错误: 源文件不存在: {source_file}")
        return
    
    if not os.path.exists('/Applications/LibreOffice.app/Contents/MacOS/soffice'):
        print("错误: 未找到LibreOffice，请先安装LibreOffice")
        return
    
    splitter = DocxSplitterV2(source_file, output_dir)
    splitter.run()


if __name__ == "__main__":
    main()
