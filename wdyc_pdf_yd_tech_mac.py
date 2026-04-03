#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word文档拆分工具 - 完全保留格式版本
直接操作Word文档的XML结构，完全保留原始格式
"""

import os
import subprocess
import shutil
import re
import glob
from docx import Document
from docx.oxml import parse_xml
from lxml import etree
import zipfile
import tempfile
try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("警告: PyPDF2未安装,将无法删除空白页")
    print("请运行: pip install PyPDF2")
    PdfReader = None
    PdfWriter = None


class DocxSplitterV2:
    def __init__(self, source_file, output_dir):
        self.source_file = source_file
        self.output_dir = output_dir
        
        # 为了解决网络驱动器权限问题 (SMB锁/元数据写入受限)，先将源文件复制到本地临时文件夹
        self.temp_workspace = tempfile.TemporaryDirectory()
        self.local_source = os.path.join(self.temp_workspace.name, "local_source.docx")
        shutil.copyfile(source_file, self.local_source)
        
        # 使用本地副本进行后续所有的搜索和提取操作
        self.doc = Document(self.local_source)
    
    def add_prefix(self, name):
        """为输出文件/文件夹名添加统一前缀"""
        return name if name.startswith("皖电云采_") else f"皖电云采_{name}"
    
    def build_output_path(self, filename, directory=None):
        """在指定目录下生成带前缀的输出路径"""
        target_dir = directory or self.output_dir
        return os.path.join(target_dir, self.add_prefix(filename))
    
    def iter_all_paragraphs(self, doc):
        """遍历文档中所有段落，包括表格里的段落"""
        for p in doc.paragraphs:
            yield p
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p
    
    def remove_chinese_heading_prefix(self, doc):
        """移除段落前的中文序号 (如 '一、'、'二、'、'十一、'等)"""
        # 允许形式：
        #   "一、标题" / "十一、标题" / "十一 、标题" / "11、标题"
        #   前面可以有空格或全角空格
        pattern = re.compile(r'^[\s\u3000]*[一二三四五六七八九十百千万零〇0-9]+[\s\u3000]*[、\.．)]')
        for para in self.iter_all_paragraphs(doc):
            if not para.text:
                continue
            
            match = pattern.match(para.text)
            if not match:
                continue
            # 直接使用匹配结束位置作为要删除的前缀长度
            prefix_len = match.end()
            consumed = 0
            for run in para.runs:
                run_text = run.text
                run_len = len(run_text)
                if consumed >= prefix_len:
                    break
                if consumed + run_len <= prefix_len:
                    run.text = ""
                    consumed += run_len
                else:
                    cut_pos = prefix_len - consumed
                    run.text = run_text[cut_pos:]
                    consumed = prefix_len
                    break
        
    def debug_print_para_info(self, idx, context=""):
        if 0 <= idx < len(self.doc.paragraphs):
            p = self.doc.paragraphs[idx]
            level = self.get_paragraph_style_level(p)
            print(f"  [DEBUG] {context} Index: {idx}, Level: {level}, Style: {p.style.name}, Text: {p.text[:50]}...")

    def is_valid_heading(self, para):
        """判断是否为有效的标题（排除目录项）"""
        text = para.text.strip()
        if not text:
            return False
            
        # 排除包含省略号的目录项
        if '...' in text or '…' in text:
            return False
            
        # 排除样式名称中包含TOC或目录的（不区分大小写）
        style_name = para.style.name.lower()
        if 'toc' in style_name or '目录' in style_name:
            return False
            
        return True

    def get_paragraph_style_level(self, paragraph):
        """获取段落的标题级别"""
        style_name = paragraph.style.name
        if 'Heading' in style_name or '标题' in style_name:
            match = re.search(r'(\d+)', style_name)
            if match:
                return int(match.group(1))
        return None
    
    def find_toc_end_index(self):
        """找到目录结束的位置（目录之后的第一个正文段落索引）"""
        w_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        toc_end = None
        in_toc_region = False
        
        for i, para in enumerate(self.doc.paragraphs):
            para_element = para._element
            text = para.text.strip()
            
            # 检查段落是否包含TOC字段指令
            has_toc_field = False
            for child in para_element.iter():
                if child.tag == f'{w_namespace}instrText':
                    if child.text and 'TOC' in child.text:
                        has_toc_field = True
                        in_toc_region = True
                        break
            
            # 检查是否是"目录"标题
            if text in ['目录', 'Table of Contents', 'TOC']:
                in_toc_region = True
            
            # 如果找到包含TOC字段的段落，记录其后的位置
            if has_toc_field:
                toc_end = i + 1
            # 如果已经过了目录区域，且找到第一个真正的正文标题（不是TOC样式），则返回
            elif in_toc_region:
                # 检查是否是真正的标题（不是TOC样式）
                style_name = para.style.name.lower()  # 转为小写
                if 'toc' not in style_name and '目录' not in style_name:
                    # 检查是否有标题级别，或者文本看起来像真正的标题（不是目录项）
                    level = self.get_paragraph_style_level(para)
                    # 如果是有级别的标题，或者文本不包含省略号（目录项通常有省略号）
                    if level is not None or ('...' not in text and '…' not in text and len(text) > 0):
                        # 确保不是目录项格式（目录项通常有页码）
                        if not re.search(r'\d+\s*$', text):  # 不以数字结尾（页码）
                            return i
        
        # 如果没有找到明确的目录结束位置，返回None（表示没有目录或目录在文档末尾）
        return toc_end
    
    def find_paragraph_indices(self, start_keyword, end_keyword=None, level=None):
        """查找段落的起始和结束索引"""
        start_idx = None
        end_idx = None
        
        for i, para in enumerate(self.doc.paragraphs):
            if start_keyword in para.text:
                start_idx = i
            elif end_keyword and end_keyword in para.text and start_idx is not None:
                end_idx = i
                break
            elif start_idx is not None and level is not None:
                para_level = self.get_paragraph_style_level(para)
                if para_level is not None and para_level <= level:
                    end_idx = i
                    break
        
        return start_idx, end_idx
    
    def find_first_level3_in_range(self, start_para_idx, end_para_idx):
        """找到指定段落范围内第一个3级标题的索引"""
        for i in range(start_para_idx, len(self.doc.paragraphs)):
            # 如果超出范围，停止
            if end_para_idx is not None and i >= end_para_idx:
                break
            
            para = self.doc.paragraphs[i]
            level = self.get_paragraph_style_level(para)
            
            # 找到第一个3级标题
            if level == 3:
                return i
        
        return None
    
    def find_level3_sections_by_parent(self, parent_para_idx, max_idx=None):
        sections = []
        parent_para = self.doc.paragraphs[parent_para_idx]
        parent_level = self.get_paragraph_style_level(parent_para)
        current_section = None
        all_level3_titles = []
        search_end = max_idx if max_idx is not None else len(self.doc.paragraphs)

        for i in range(parent_para_idx + 1, search_end):
            para = self.doc.paragraphs[i]
            level = self.get_paragraph_style_level(para)

            if level is not None and parent_level is not None and level <= parent_level:
                if current_section:
                    current_section['end_idx'] = i
                    sections.append(current_section)
                break

            if level == 3:
                title_text = para.text.strip()
                all_level3_titles.append((i, title_text))

                should_skip = False

                skip_keywords = ['表', '图', '附件', '目录', 'TOC', '...', '简历', '身份证', '毕业证', '资格证书', '职称', '奖项', '荣誉', '清单', '报价', '明细', '工作量', '图纸', '方案', '报告', '计划']
                for keyword in skip_keywords:
                    if keyword in title_text:
                        should_skip = True
                        break

                if len(title_text) < 3:
                    should_skip = True

                if re.match(r'^[\d\s、\.．)]+$', title_text):
                    should_skip = True

                if re.match(r'^\d+\.\s*[\u4e00-\u9fa5]+\s*[（(][^）)]+[）)]', title_text):
                    should_skip = True

                position_keywords = ['项目经理', '调度主管', '技术负责人', '质量负责人', '安全负责人', '财务负责人', '采购负责人', '销售负责人', '运营负责人', '主管', '经理', '总监', '工程师', '专员', '助理']
                for keyword in position_keywords:
                    if keyword in title_text and '（' in title_text:
                        should_skip = True
                        break

                try:
                    in_table = bool(para._element.xpath('ancestor::w:tbl', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
                except Exception:
                    in_table = False
                if in_table:
                    should_skip = True

                project_keywords = ['合同', '协议', '项目']
                has_project_keyword = any(keyword in title_text for keyword in project_keywords)

                content_keywords = ['合同编号', '合同名称', '项目名称', '签订日期', '签订时间', '采购单位', '建设单位', '服务单位', '中标通知书', '验收', '履约', '项目地址', '合同金额', '标的金额']
                found_content = False
                for j in range(i + 1, search_end):
                    p2 = self.doc.paragraphs[j]
                    lv2 = self.get_paragraph_style_level(p2)
                    if lv2 is not None and lv2 <= 3:
                        break
                    text2 = p2.text.strip()
                    if not text2:
                        continue
                    if any(k in text2 for k in content_keywords):
                        found_content = True
                        break

                if should_skip:
                    continue

                if current_section:
                    current_section['end_idx'] = i
                    sections.append(current_section)

                current_section = {
                    'title': title_text,
                    'start_idx': i,
                    'end_idx': None
                }

        # 添加最后一个section
        if current_section:
            if not current_section['end_idx']:
                # 如果搜索到了文档末尾，则将end_idx设为None，以包含后续的所有内容（如表格）
                if search_end == len(self.doc.paragraphs):
                    current_section['end_idx'] = None
                else:
                    current_section['end_idx'] = search_end
            sections.append(current_section)

        # 调试信息：显示所有找到的3级标题
        print(f"  找到 {len(sections)} 个项目")

        return sections
    
    def find_toc_region(self, doc):
        """找到目录区域的范围（起始和结束段落索引）"""
        w_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        toc_start = None
        toc_end = None
        
        for i, para in enumerate(doc.paragraphs):
            para_element = para._element
            text = para.text.strip()
            
            # 检查是否是"目录"标题
            if text in ['目录', 'Table of Contents', 'TOC']:
                if toc_start is None:
                    toc_start = i
            
            # 检查是否包含TOC字段指令
            has_toc_field = False
            for child in para_element.iter():
                if child.tag == f'{w_namespace}instrText':
                    if child.text and 'TOC' in child.text:
                        has_toc_field = True
                        if toc_start is None:
                            toc_start = i
                        break
            
            # 如果找到了目录开始，继续查找目录结束
            if toc_start is not None:
                # 检查是否是目录项（包含省略号）
                is_toc_item = '...' in text or '…' in text
                # 检查样式是否为TOC样式
                style_name = para.style.name.lower()
                is_toc_style = 'toc' in style_name
                # 检查是否像目录项（有页码）
                has_page_number = bool(re.search(r'\d+\s*$', text))
                
                # 如果这个段落看起来像目录项，更新toc_end
                if is_toc_item or (is_toc_style and has_page_number) or has_toc_field:
                    toc_end = i + 1
                # 如果已经过了目录区域，找到第一个真正的正文标题，则目录结束
                elif toc_end is not None:
                    # 检查是否是真正的标题（不是TOC样式）
                    if 'toc' not in style_name and '目录' not in style_name:
                        level = self.get_paragraph_style_level(para)
                        if level is not None:
                            # 找到了真正的正文标题，目录区域结束
                            break
        
        return toc_start, toc_end
    
    def extract_docx_by_paragraph_range(self, start_para_idx, end_para_idx, output_path):
        """通过复制整个docx文件并删除不需要的段落来提取内容，完全保留格式"""
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            # 复制本地源文件副本到处理位置 (避免直接操作网络驱动器)
            temp_docx = os.path.join(temp_dir, 'temp.docx')
            shutil.copyfile(self.local_source, temp_docx)
            
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
                    
                    # 检查是否在提取范围外
                    if para_count < start_para_idx or (end_para_idx and para_count >= end_para_idx):
                        should_remove = True
                        # 如果段落在范围外，检查是否是目录项，如果是则删除
                        if para_count < len(para_styles):
                            style_name = para_styles[para_count]
                            if 'toc' in style_name:
                                should_remove = True
                    else:
                        # 段落在提取范围内，需要保留
                        # 但如果是明确的目录项，仍然要删除
                        
                        # 先获取段落文本
                        para_text = ''
                        for child in element.iter():
                            if child.tag == f'{w_namespace}t':
                                if child.text:
                                    para_text += child.text
                        
                        # 检查是否包含TOC字段指令（这是最明确的目录标识）
                        has_toc_field = False
                        for child in element.iter():
                            if child.tag == f'{w_namespace}instrText':
                                if child.text and 'TOC' in child.text:
                                    has_toc_field = True
                                    break
                        
                        # 如果包含TOC字段指令，删除（这是目录字段本身）
                        if has_toc_field:
                            should_remove = True
                        # 如果段落文本完全是"目录"，删除
                        elif para_text.strip() in ['目录', 'Table of Contents', 'TOC']:
                            should_remove = True
                        # 如果段落包含省略号且样式是TOC样式，删除（目录项）
                        elif ('...' in para_text or '…' in para_text) and para_count < len(para_styles):
                            style_name = para_styles[para_count].lower()  # 转为小写
                            if 'toc' in style_name:
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
                
                # 处理其他元素 (如 sdt - 可能是目录)
                elif element.tag.endswith('sectPr'):
                    # 保留节属性，以免破坏页面布局
                    pass
                else:
                    # 对于其他元素(如sdt)，如果当前位置不在范围内，则删除
                    if para_count < start_para_idx or (end_para_idx and para_count >= end_para_idx):
                        elements_to_remove.append(element)
            
            # 删除不需要的元素
            for element in elements_to_remove:
                body.remove(element)
            
            # 去除标题中的中文序号
            self.remove_chinese_heading_prefix(doc)
            
            # 保存修改后的文档
            doc.save(output_path)
            
            # 展开字段为普通文本，避免显示"Error: Reference source not found"
            self.expand_fields_in_docx(output_path)
    
    def extract_docx_by_multiple_ranges(self, ranges, output_path):
        """通过复制整个docx文件并只保留指定范围内的段落来提取内容"""
        # ranges 是一个列表，包含 (start_idx, end_idx) 元组
        
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            # 复制本地源文件副本到处理位置 (避免直接操作网络驱动器)
            temp_docx = os.path.join(temp_dir, 'temp.docx')
            shutil.copyfile(self.local_source, temp_docx)
            
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
                    should_remove = True
                    
                    # 检查是否在任何一个范围内
                    for start_idx, end_idx in ranges:
                        if para_count >= start_idx and (end_idx is None or para_count < end_idx):
                            should_remove = False
                            break
                    
                    # 如果在范围内，还需要检查是否是目录或TOC
                    if not should_remove:
                        # 先获取段落文本
                        para_text = ''
                        for child in element.iter():
                            if child.tag == f'{w_namespace}t':
                                if child.text:
                                    para_text += child.text
                        
                        # 检查是否包含TOC字段指令（这是最明确的目录标识）
                        has_toc_field = False
                        for child in element.iter():
                            if child.tag == f'{w_namespace}instrText':
                                if child.text and 'TOC' in child.text:
                                    has_toc_field = True
                                    break
                        
                        # 如果包含TOC字段指令，删除（这是目录字段本身）
                        if has_toc_field:
                            should_remove = True
                        # 如果段落文本完全是"目录"，删除
                        elif para_text.strip() in ['目录', 'Table of Contents', 'TOC']:
                            should_remove = True
                        # 如果段落包含省略号且样式是TOC样式，删除（目录项）
                        elif ('...' in para_text or '…' in para_text) and para_count < len(para_styles):
                            style_name = para_styles[para_count]
                            if 'toc' in style_name:
                                should_remove = True
                    
                    if should_remove:
                        elements_to_remove.append(element)
                    
                    para_count += 1
                    
                # 如果是表格，检查它是否在范围内
                elif element.tag.endswith('tbl'):
                    # 表格比较复杂，我们需要检查它前后的段落位置
                    # 简单起见，如果当前段落计数在任何一个范围内，就保留表格
                    should_remove = True
                    for start_idx, end_idx in ranges:
                        if para_count >= start_idx and (end_idx is None or para_count < end_idx):
                            should_remove = False
                            break
                    
                    if should_remove:
                        elements_to_remove.append(element)
                
                # 处理其他元素 (可能是目录)
                elif element.tag.endswith('sectPr'):
                    # 保留节属性，以免破坏页面布局
                    pass
                else:
                    # 对于其他元素(如sdt)，如果当前位置不在范围内，则删除
                    # 这有助于删除位于文档开头的TOC控件
                    should_remove = True
                    for start_idx, end_idx in ranges:
                        if para_count >= start_idx and (end_idx is None or para_count < end_idx):
                            should_remove = False
                            break
                    
                    # 如果在范围内，但包含TOC字段，仍然删除
                    if not should_remove:
                        for child in element.iter():
                            if child.tag == f'{w_namespace}instrText':
                                if child.text and 'TOC' in child.text:
                                    should_remove = True
                                    break
                    
                    if should_remove:
                        elements_to_remove.append(element)
            
            # 删除不需要的元素
            for element in elements_to_remove:
                body.remove(element)
            
            # 去除标题中的中文序号
            self.remove_chinese_heading_prefix(doc)
            
            # 保存修改后的文档
            doc.save(output_path)
            
            # 展开字段为普通文本，避免显示"Error: Reference source not found"
            self.expand_fields_in_docx(output_path)
    

    
    def expand_fields_in_docx(self, docx_path):
        """展开Word文档中的字段为普通文本，避免显示Error: Reference source not found"""
        try:
            import zipfile
            from xml.etree import ElementTree as ET
            
            # Word命名空间
            w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            w_ns_prefixed = '{' + w_ns + '}'
            
            # 打开docx文件（实际上是一个zip文件）
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                # 读取主文档XML
                document_xml = zip_ref.read('word/document.xml')
            
            # 解析XML
            root = ET.fromstring(document_xml)
            
            # 注册命名空间
            ET.register_namespace('w', w_ns)
            
            # 定义命名空间字典
            namespaces = {'w': w_ns}
            
            # 处理简单字段（fldSimple）- 这些是内联字段
            field_simples = root.findall('.//w:fldSimple', namespaces)
            for field_simple in field_simples:
                # 获取字段指令
                instr = field_simple.find('w:instrText', namespaces)
                if instr is not None:
                    instr_text = instr.text or ''
                    # 查找字段结果文本（在w:r/w:t中）
                    result_texts = field_simple.findall('.//w:t', namespaces)
                    result_text = ''.join([(t.text or '') for t in result_texts])
                    
                    # 如果结果文本为空或者是错误信息，用占位符替换
                    if not result_text or 'Error' in result_text or 'not found' in result_text:
                        result_text = ''
                    
                    # 用纯文本运行替换整个字段
                    parent = field_simple.getparent()
                    if parent is not None:
                        # 创建文本运行
                        new_r = ET.Element(w_ns_prefixed + 'r')
                        new_t = ET.SubElement(new_r, w_ns_prefixed + 't')
                        new_t.text = result_text if result_text else ''
                        # 替换字段
                        parent.replace(field_simple, new_r)
            
            # 处理复杂字段（使用fldChar的字段）
            # 这些字段由fldChar（开始）、instrText（指令）和fldChar（结束）组成
            # 使用lxml重新解析以便使用XPath（代码开头已导入lxml.etree）
            lxml_root = etree.fromstring(document_xml)
            lxml_ns = {'w': w_ns}
            
            # 查找所有字段开始
            field_begins = lxml_root.xpath('.//w:fldChar[@w:fldCharType="begin"]', namespaces=lxml_ns)
            
            for field_begin in field_begins:
                # 找到字段开始的运行和段落
                run = field_begin.getparent()
                if run is None or not run.tag.endswith('}r'):
                    continue
                    
                para = run.getparent()
                if para is None or not para.tag.endswith('}p'):
                    continue
                
                # 查找字段结束标记（在同一段落中）
                field_end = None
                result_text = ''
                runs_to_remove = [run]
                
                # 从字段开始后查找
                found_begin = False
                for sibling in para:
                    if sibling == run:
                        found_begin = True
                        continue
                    if not found_begin:
                        continue
                    
                    if sibling.tag.endswith('}r'):
                        # 检查是否是字段结束
                        fld_char_end = sibling.xpath('.//w:fldChar[@w:fldCharType="end"]', namespaces=lxml_ns)
                        if fld_char_end:
                            field_end = sibling
                            runs_to_remove.append(sibling)
                            break
                        
                        # 收集结果文本（在字段结束之前）
                        texts = sibling.xpath('.//w:t', namespaces=lxml_ns)
                        for text_elem in texts:
                            if text_elem.text:
                                result_text += text_elem.text
                        
                        # 检查是否有指令文本，如果有则标记为要删除
                        instr = sibling.xpath('.//w:instrText', namespaces=lxml_ns)
                        if instr:
                            runs_to_remove.append(sibling)
                        elif not field_end:
                            # 如果还没有找到结束标记，可能是字段结果的一部分
                            runs_to_remove.append(sibling)
                
                # 如果找到了字段结构，替换它
                if field_end is not None or len(runs_to_remove) > 1:
                    # 如果结果文本包含错误信息，清空它
                    if 'Error' in result_text or 'not found' in result_text:
                        result_text = ''
                    
                    # 创建新的文本运行替换第一个字段运行
                    new_r = etree.Element(w_ns_prefixed + 'r')
                    new_t = etree.SubElement(new_r, w_ns_prefixed + 't')
                    new_t.text = result_text
                    
                    # 替换第一个运行
                    para.replace(runs_to_remove[0], new_r)
                    
                    # 删除其他字段相关的运行
                    for run_to_remove in runs_to_remove[1:]:
                        if run_to_remove.getparent() == para:
                            para.remove(run_to_remove)
            
            # 使用lxml处理后的结果
            modified_xml = etree.tostring(lxml_root, encoding='utf-8', xml_declaration=True)
            
            # 重新创建zip文件（不能直接修改，需要重新创建）
            temp_docx = docx_path + '.tmp'
            with zipfile.ZipFile(docx_path, 'r') as zip_in:
                with zipfile.ZipFile(temp_docx, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for item in zip_in.infolist():
                        if item.filename == 'word/document.xml':
                            zip_out.writestr(item, modified_xml)
                        else:
                            zip_out.writestr(item, zip_in.read(item.filename))
            
            # 替换原文件
            shutil.move(temp_docx, docx_path)
                
        except Exception as e:
            # 如果展开失败，不影响主流程
            print(f"  警告: 展开字段时出错: {e}")
            pass
    
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
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            # LibreOffice会生成与输入文件同名的PDF
            generated_pdf = docx_path.replace('.docx', '.pdf')
            
            if os.path.exists(generated_pdf):
                if generated_pdf != pdf_path:
                    shutil.move(generated_pdf, pdf_path)
                print(f"✓ 已生成: {os.path.basename(pdf_path)}")
                return True
            else:
                print(f"✗ 转换失败: {os.path.basename(pdf_path)}")
                return False
                
        except subprocess.TimeoutExpired:
            print(f"✗ 转换超时: {os.path.basename(pdf_path)}")
            return False
        except Exception as e:
            print(f"✗ 转换出错 {os.path.basename(pdf_path)}: {e}")
            return False
    
    def is_blank_page(self, page):
        """检测PDF页面是否为空白页"""
        try:
            # 提取页面文本
            text = page.extract_text().strip()
            
            # 检查页面内容流，看是否真正使用了图片
            has_displayed_image = False
            if '/Contents' in page:
                try:
                    # 获取页面内容流
                    contents = page['/Contents']
                    if hasattr(contents, 'get_data'):
                        content_data = contents.get_data().decode('latin-1', errors='ignore')
                    else:
                        # 如果是数组，合并所有内容流
                        content_data = ''
                        for content in contents:
                            if hasattr(content, 'get_data'):
                                content_data += content.get_data().decode('latin-1', errors='ignore')
                    
                    # 检查内容流中是否有 "Do" 操作符（绘制XObject/图片的命令）
                    if '/Im' in content_data and ' Do' in content_data:
                        has_displayed_image = True
                except:
                    pass
            
            # 如果页面有实际显示的图片，不是空白页
            if has_displayed_image:
                return False
            
            # 如果文本长度很短(少于5个字符),认为是空白页
            if len(text) < 5:
                return True
            
            # 检查是否只包含空白字符
            if not text or text.isspace():
                print(f"  [DEBUG] Page is blank (empty text)")
                return True
                
            return False
        except Exception as e:
            # 如果检测过程出错，保守起见不删除该页
            return False
    
    def remove_blank_pages(self, pdf_path):
        """删除PDF中的空白页"""
        if not PdfReader or not PdfWriter:
            print("  PyPDF2未安装,跳过删除空白页")
            return False
            
        try:
            # 读取PDF
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            
            total_pages = len(reader.pages)
            blank_pages = []
            
            # 检查每一页
            for i, page in enumerate(reader.pages):
                if self.is_blank_page(page):
                    blank_pages.append(i + 1)
                else:
                    writer.add_page(page)
            
            # 如果有空白页被删除,保存新PDF
            if blank_pages:
                # 创建临时文件
                temp_pdf = pdf_path + '.tmp'
                with open(temp_pdf, 'wb') as f:
                    writer.write(f)
                
                # 替换原文件
                shutil.move(temp_pdf, pdf_path)
                print(f"  已删除 {len(blank_pages)} 个空白页 (总共 {total_pages} 页): 第 {', '.join(map(str, blank_pages))} 页")
                return True
            else:
                print(f"  未发现空白页 (总共 {total_pages} 页)")
                return False
                
        except Exception as e:
            print(f"  删除空白页时出错: {e}")
            return False
    
    def remove_last_page(self, pdf_path):
        """删除PDF的最后一页"""
        if not PdfReader or not PdfWriter:
            print("  PyPDF2未安装,跳过删除最后一页")
            return False
            
        try:
            if not os.path.exists(pdf_path):
                return False
                
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            
            if total_pages <= 1:
                print(f"  PDF只有 {total_pages} 页，无法删除最后一页")
                return False
            
            writer = PdfWriter()
            
            # 复制除最后一页外的所有页
            for i in range(total_pages - 1):
                writer.add_page(reader.pages[i])
            
            # 创建临时文件
            temp_pdf = pdf_path + '.tmp'
            with open(temp_pdf, 'wb') as f:
                writer.write(f)
            
            # 替换原文件
            shutil.move(temp_pdf, pdf_path)
            print(f"  已删除最后一页 (原共 {total_pages} 页，现 {total_pages - 1} 页)")
            return True
            
        except Exception as e:
            print(f"  删除最后一页时出错: {e}")
            return False
    
    def save_section_as_pdf(self, start_idx, end_idx, output_path):
        """保存段落范围为PDF"""
        temp_docx = output_path.replace('.pdf', '_temp.docx')
        
        # 提取内容到临时docx
        self.extract_docx_by_paragraph_range(start_idx, end_idx, temp_docx)
        
        # 转换为PDF
        success = self.convert_docx_to_pdf(temp_docx, output_path)
        
        # 如果转换成功,删除空白页
        if success and os.path.exists(output_path):
            self.remove_blank_pages(output_path)
        
        # 删除临时文件
        if os.path.exists(temp_docx):
            os.remove(temp_docx)
        
    def save_sections_as_pdf(self, ranges, output_path):
        """保存多个段落范围为一个PDF"""
        temp_docx = output_path.replace('.pdf', '_temp.docx')
        
        # 提取内容到临时docx
        self.extract_docx_by_multiple_ranges(ranges, temp_docx)
        
        # 转换为PDF
        success = self.convert_docx_to_pdf(temp_docx, output_path)
        
        # 如果转换成功,删除空白页
        if success and os.path.exists(output_path):
            self.remove_blank_pages(output_path)
        
        # 删除临时文件
        if os.path.exists(temp_docx):
            os.remove(temp_docx)
        
        return success
    
    def split_related_materials(self):
        """拆分"相关证明材料按项目逐一列明"下的3级目录"""
        print("\n1. 拆分相关证明材料...")
        output_folder = os.path.join(self.output_dir, self.add_prefix("企业类似项目业绩和实施经验"))
        os.makedirs(output_folder, exist_ok=True)
        
        parent_idx = None
        for i, para in enumerate(self.doc.paragraphs):
            if ("相关证明材料按项目" in para.text or "业绩文件" in para.text) and self.is_valid_heading(para):
                parent_idx = i
                break
        
        if parent_idx is None:
            print("  未找到'（二）相关证明材料按项目逐一列明' 或 '业绩文件'")
            return
            
        # 尝试找到下一个大章节的开始位置，作为搜索的截止点
        next_section_idx = None
        
        # 1. 搜索常见的下一章节标题
        next_keywords = ["拟委任", "项目团队", "技术方案", "服务方案", "（三）", "三、"]
        
        for i in range(parent_idx + 1, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            text = para.text.strip()
            level = self.get_paragraph_style_level(para)
            
            # 如果遇到同级或更高级别的标题
            parent_level = self.get_paragraph_style_level(self.doc.paragraphs[parent_idx])
            if level is not None and parent_level is not None and level <= parent_level:
                next_section_idx = i
                break
                
            # 如果遇到关键词匹配的段落（可能是标题但样式不对）
            for kw in next_keywords:
                if kw in text and len(text) < 50:
                    # 再次确认不是普通正文
                    if level is not None or any(run.bold for run in para.runs):
                        next_section_idx = i
                        break
            if next_section_idx:
                break
        
        sections = self.find_level3_sections_by_parent(parent_idx, max_idx=next_section_idx)
        print(f"  找到 {len(sections)} 个项目")
        
        for idx, section in enumerate(sections, 1):
            title = section['title']
            safe_title = re.sub(r'[\\/:*?"<>|]', '_', title)
            output_path = self.build_output_path(f"{idx:02d}_{safe_title}.pdf", output_folder)
            self.save_section_as_pdf(section['start_idx'], section['end_idx'], output_path)
    
    def split_management_team(self):
        """拆分"拟委任的主要人员汇总表（管理团队）"部分"""
        print("\n2. 拆分管理团队汇总表...")
        
        # 先找到目录结束的位置，只在此之后查找
        toc_end = self.find_toc_end_index()
        start_search_idx = toc_end if toc_end is not None else 0
        
        if toc_end is not None:
            print(f"  目录结束位置: 段落 {toc_end}")
        
        parent_idx = None
        for i in range(start_search_idx, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            text = para.text.strip()
            # 过滤掉可能存在于目录中的条目
            if ("拟委任的主要人员汇总表" in text or "项目团队情况" in text) and self.is_valid_heading(para):
                # 额外检查：确保不是目录项（检查是否包含TOC字段）
                w_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                para_element = para._element
                has_toc_field = False
                for child in para_element.iter():
                    if child.tag == f'{w_namespace}instrText':
                        if child.text and 'TOC' in child.text:
                            has_toc_field = True
                            break
                
                # 如果包含TOC字段，跳过（这是目录项）
                if has_toc_field:
                    continue
                
                parent_idx = i
                break
        
        if parent_idx is None:
            print("  未找到'拟委任的主要人员汇总表' 或 '项目团队情况'章节")
            return
        
        # 查找结束位置：下一个同级或更高级别的标题
        end_idx = None
        current_level = self.get_paragraph_style_level(self.doc.paragraphs[parent_idx])
        
        for i in range(parent_idx + 1, len(self.doc.paragraphs)):
            level = self.get_paragraph_style_level(self.doc.paragraphs[i])
            if level is not None:
                # 如果起始段落是标题 (有level)
                if current_level is not None:
                    # 如果找到同级或更高级别的标题，则停止
                    if level <= current_level:
                        end_idx = i
                        break
                # 如果起始段落不是标准标题 (没有level)
                else:
                    # 假定它是一个重要的部分，直到下一个1级或2级标题都算作其内容
                    if level <= 2:
                        end_idx = i
                        break

        # 如果没有找到结束标题，则到文档末尾
        if end_idx is None:
            end_idx = None # 使用None表示包含直到文件末尾的所有内容（包括表格）
            
        print(f"  找到 '拟委任的主要人员汇总表' 范围: 段落 {parent_idx} - {end_idx}")

        if end_idx is not None and parent_idx >= end_idx:
            print("  警告: 找到的范围为空，可能未能正确识别章节结束位置。")
            return
        
        # 查找该范围内第一个3级标题的位置（人员简历部分）
        first_level3_idx = self.find_first_level3_in_range(parent_idx, end_idx)
        
        if first_level3_idx is not None:
            print(f"  找到第一个3级标题位置: 段落 {first_level3_idx}")
            # 拆分成两个文件：表格部分和简历部分
            # Part 1: 从开始到第一个3级标题之前（表格部分）
            output_path_table = self.build_output_path("项目团队情况.pdf")
            self.save_section_as_pdf(parent_idx, first_level3_idx, output_path_table)
            
            # Part 2: 从第一个3级标题到结束（简历部分）
            if end_idx is None or first_level3_idx < end_idx:
                output_path_resume = self.build_output_path("主要人员简历表及证明文件.pdf")
                self.save_section_as_pdf(first_level3_idx, end_idx, output_path_resume)
            else:
                print("  3级标题后没有内容，跳过简历部分")
        else:
            print("  未找到3级标题，生成完整PDF")
            # 如果没有找到3级标题，生成完整的PDF
            output_path = self.build_output_path("拟委任的主要人员汇总表.pdf")
            self.save_section_as_pdf(parent_idx, end_idx, output_path)
    
    def split_chapters_4_to_6(self):
        """拆分第四到第六大章"""
        print("\n3. 拆分第四到第六章...")
        
        start_idx = None
        end_idx = None
        
        for i, para in enumerate(self.doc.paragraphs):
            level = self.get_paragraph_style_level(para)
            text = para.text.strip()
            
            # 跳过无效标题（如目录）
            if not self.is_valid_heading(para):
                continue
            
            if level == 1 and ('四、' in text or '第四' in text):
                start_idx = i
            
            # 如果已经找到了开始位置，且当前是1级标题
            if start_idx and level == 1:
                # 如果标题不是四、五、六章，那么就是结束位置（比如第七章，或者其他一级标题）
                if not any(x in text for x in ['四、', '第四', '五、', '第五', '六、', '第六']):
                    end_idx = i
                    break
        
        if start_idx:
            output_path = self.build_output_path("工作方案及服务承诺.pdf")
            self.save_section_as_pdf(start_idx, end_idx, output_path)
        else:
            print("  未找到第四章")
    
    def split_patents(self):
        """拆分专利章节到一个PDF文件"""
        print("\n4. 拆分专利...")
        
        # 先找到目录结束的位置，只在此之后查找
        toc_end = self.find_toc_end_index()
        start_search_idx = toc_end if toc_end is not None else 0
        
        if toc_end is not None:
            print(f"  目录结束位置: 段落 {toc_end}")
        
        parent_idx = None
        parent_level = None
        for i in range(start_search_idx, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            text = para.text.strip()
            # 增加对 "专利信息" 的匹配
            if ("专利" in text and ("数量" in text or "情况" in text or "列表" in text or "信息" in text)) and self.is_valid_heading(para):
                parent_idx = i
                parent_level = self.get_paragraph_style_level(para)
                self.debug_print_para_info(i, f"Found Patent Start (Level {parent_level})")
                break
        
        if parent_idx is None:
            print("  未找到包含'专利'关键字的章节")
            return
        
        # 找到专利章节的结束位置 (下一个同级或更高级别的标题)
        next_parent_idx = len(self.doc.paragraphs)
        if parent_level is not None:
            for i in range(parent_idx + 1, len(self.doc.paragraphs)):
                level = self.get_paragraph_style_level(self.doc.paragraphs[i])
                # 如果当前是1级标题，则直到下一个1级标题才结束（从而包含其中的所有2级标题）
                if level is not None and level <= parent_level:
                    next_parent_idx = i
                    break
        else:
            # 如果父标题没有级别，找下一个2级或1级标题
            for i in range(parent_idx + 1, len(self.doc.paragraphs)):
                level = self.get_paragraph_style_level(self.doc.paragraphs[i])
                if level is not None and level <= 2:
                    next_parent_idx = i
                    break

        # 始终作为一个整体导出
        output_path = self.build_output_path("专利.pdf")
        end_idx = next_parent_idx if next_parent_idx < len(self.doc.paragraphs) else None
        self.save_section_as_pdf(parent_idx, end_idx, output_path)
    
    def split_performance_evaluations(self):
        """拆分绩效评价章节到一个PDF文件"""
        print("\n5. 拆分绩效评价...")
        
        # 先找到目录结束的位置，只在此之后查找
        toc_end = self.find_toc_end_index()
        start_search_idx = toc_end if toc_end is not None else 0
        
        if toc_end is not None:
            print(f"  目录结束位置: 段落 {toc_end}")
        
        # 查找"绩效评价"的起始位置
        parent_idx = None
        for i in range(start_search_idx, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            text = para.text.strip()
            if ("绩效评价" in text or "履约评价" in text) and self.is_valid_heading(para):
                parent_idx = i
                self.debug_print_para_info(i, "Found Performance Start")
                break
        
        if parent_idx is None:
            print("  未找到'绩效评价'")
            return
        
        # 找到绩效评价章节的结束位置（下一个2级或更高级别的标题）
        end_idx = None
        for i in range(parent_idx + 1, len(self.doc.paragraphs)):
            level = self.get_paragraph_style_level(self.doc.paragraphs[i])
            if level is not None and level <= 2:
                end_idx = i
                self.debug_print_para_info(i, "Found Performance End")
                break
        
        # 如果没有找到结束位置，使用文档末尾
        if end_idx is None:
            end_idx = None
        
        print(f"  找到'绩效评价'范围: {parent_idx} - {end_idx}")
        
        output_path = self.build_output_path("绩效评价.pdf")
        self.save_section_as_pdf(parent_idx, end_idx, output_path)
        
        # 删除最后一页（该页内容错误）
        self.remove_last_page(output_path)
    
    def split_additional_sections(self):
        """拆分"技术偏差表"、"研发团队人员汇总表"、"公司资质"到一个PDF"""
        print("\n6. 拆分技术偏差表+研发团队+公司资质...")
        
        # 先找到目录结束的位置，只在此之后查找
        toc_end = self.find_toc_end_index()
        start_search_idx = toc_end if toc_end is not None else 0
        
        if toc_end is not None:
            print(f"  目录结束位置: 段落 {toc_end}")
        
        ranges = []
        
        # 1. 查找"技术偏差表"范围
        tech_start = None
        tech_end = None
        for i in range(start_search_idx, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            if "技术偏差表" in para.text and self.is_valid_heading(para):
                tech_start = i # 包含标题
                break
        
        if tech_start is not None:
            # 找到下一个1级标题作为结束
            for i in range(tech_start + 1, len(self.doc.paragraphs)):
                level = self.get_paragraph_style_level(self.doc.paragraphs[i])
                if level == 1:
                    tech_end = i
                    break
            if tech_end is None:
                tech_end = None
            ranges.append((tech_start, tech_end))
            print(f"  找到'技术偏差表': {tech_start} - {tech_end}")
        else:
            print("  未找到'技术偏差表'")

        # 2. 查找"研发团队人员汇总表"范围
        team_start = None
        team_end = None
        for i in range(start_search_idx, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            if "研发团队人员汇总表" in para.text and self.is_valid_heading(para):
                team_start = i # 包含标题
                break
        
        if team_start is not None:
            # 找到下一个1级标题或"完备认证体系"作为结束
            for i in range(team_start + 1, len(self.doc.paragraphs)):
                para = self.doc.paragraphs[i]
                level = self.get_paragraph_style_level(para)
                # 遇到1级标题或特定文本时结束
                if level == 1 or "完备认证体系" in para.text:
                    team_end = i
                    break
            if team_end is None:
                team_end = None
            ranges.append((team_start, team_end))
            print(f"  找到'研发团队人员汇总表': {team_start} - {team_end}")
        else:
            print("  未找到'研发团队人员汇总表'")
            
        # 3. 查找"公司资质"范围
        qual_start = None
        qual_end = None
        for i in range(start_search_idx, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            if "公司资质" in para.text and self.is_valid_heading(para):
                qual_start = i # 包含标题
                break
        
        if qual_start is not None:
            # 找到下一个1级标题或"2．完备认证体系，夯实业务根基"作为结束
            for i in range(qual_start + 1, len(self.doc.paragraphs)):
                para = self.doc.paragraphs[i]
                level = self.get_paragraph_style_level(para)
                # 遇到1级标题或特定文本时结束
                if level == 1 or "完备认证体系" in para.text:
                    qual_end = i
                    break
            if qual_end is None:
                qual_end = None
            ranges.append((qual_start, qual_end))
            print(f"  找到'公司资质': {qual_start} - {qual_end}")
        else:
            print("  未找到'公司资质'")
        
        if not ranges:
            print("  未找到任何目标章节")
            return
            
        output_path = self.build_output_path("评标办法所涉及的其他.pdf")
        self.save_sections_as_pdf(ranges, output_path)
    
    # 注意: 此方法已被弃用，现在在split_management_team()中直接按表格边界拆分
    # def split_management_team_pdf_further(self):
    #     """将"拟委任的主要人员汇总表.pdf"拆分为Part1(前2页)和Part2(剩余页)"""
    #     print("\n7. 进一步拆分管理团队汇总表PDF...")
    #     
    #     pdf_path = self.build_output_path("拟委任的主要人员汇总表.pdf")
    #     if not os.path.exists(pdf_path):
    #         print(f"  错误: 文件不存在: {pdf_path}")
    #         return
    # 
    #     if not PdfReader or not PdfWriter:
    #         print("  PyPDF2未安装,跳过进一步拆分")
    #         return
    # 
    #     try:
    #         reader = PdfReader(pdf_path)
    #         total_pages = len(reader.pages)
    #         
    #         if total_pages < 2:
    #             print(f"  警告: 文件只有 {total_pages} 页，无法拆分前2页")
    #             # Fallback: copy to table pdf, skip resume pdf
    #             output1 = self.build_output_path("拟委任的主要人员汇总表_人员表格.pdf")
    #             shutil.copy2(pdf_path, output1)
    #             print(f"  ✓ 已生成: {os.path.basename(output1)} (完整文件)")
    #             return
    # 
    #         # 创建两个Writer对象
    #         writer1 = PdfWriter()
    #         writer2 = PdfWriter()
    # 
    #         # Part 1: 前2页
    #         for i in range(min(2, total_pages)):
    #             writer1.add_page(reader.pages[i])
    # 
    #         # Part 2: 剩余页
    #         for i in range(2, total_pages):
    #             writer2.add_page(reader.pages[i])
    # 
    #         # 保存文件
    #         output1 = self.build_output_path("拟委任的主要人员汇总表_人员表格.pdf")
    #         output2 = self.build_output_path("拟委任的主要人员汇总表_人员简历.pdf")
    # 
    #         with open(output1, "wb") as f1:
    #             writer1.write(f1)
    #         print(f"  ✓ 已生成: {os.path.basename(output1)} (共 {len(writer1.pages)} 页)")
    # 
    #         if len(writer2.pages) > 0:
    #             with open(output2, "wb") as f2:
    #                 writer2.write(f2)
    #             print(f"  ✓ 已生成: {os.path.basename(output2)} (共 {len(writer2.pages)} 页)")
    #         else:
    #             print("  没有剩余页面，未生成Part2")
    #             
    #         # 可选：删除原文件，或者保留
    #         os.remove(pdf_path) 
    #         
    #     except Exception as e:
    #         print(f"  拆分PDF时出错: {e}")

    def run(self):
        """执行所有拆分任务"""
        print("=" * 80)
        print(f"开始拆分文档: {os.path.basename(self.source_file)}")
        print("=" * 80)
        
        self.split_related_materials()
        self.split_management_team()
        self.split_chapters_4_to_6()
        self.split_patents()
        self.split_performance_evaluations()
        self.split_additional_sections()
        
        print("\n" + "=" * 80)
        print(f"文件 {os.path.basename(self.source_file)} 拆分完成！")
        print("=" * 80)
        
        # 清理本地临时工作区
        try:
            self.temp_workspace.cleanup()
        except:
            pass


def main():
    # ***************************************************************
    # 请输入要扫描的根目录
    # ***************************************************************
    source_dir = input("请输入要扫描的根目录路径: ")
    # 示例路径，用户可以直接复制粘贴到输入框中
    # source_dir = "/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/国网安徽电力2025年第二次服务区域联合授权竞争性谈判采购-2025-12-02-AM6：00-ECP/零星服务-合肥公司/包15_完整采购文件_78655757526447548"
    # source_dir = "/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/国网安徽电力2025年第二次服务区域联合授权竞争性谈判采购-2025-12-02-AM6：00-ECP"
    # source_dir = "/Volumes/zhuxl 共享给我/1、25年年底投标汇总8888/国网安徽电力2025年第二次服务区域联合授权竞争性谈判采购-2025-12-02-AM6：00-ECP/零星服务-芜湖公司/包23_完整采购文件_78656978500898585/"
    
    

    if not os.path.exists(source_dir):
        print(f"错误: 目录不存在: {source_dir}")
        return
    
    if not os.path.exists('/Applications/LibreOffice.app/Contents/MacOS/soffice'):
        print("错误: 未找到LibreOffice，请先安装LibreOffice")
        print("下载地址: https://www.libreoffice.org/download/download/")
        return

    # 递归查找目录及其子目录下包含“技术文件”的docx文件
    docx_files = []
    for root, _, files in os.walk(source_dir):
        for filename in files:
            if not filename.lower().endswith(".docx"):
                continue
            if "技术文件" not in filename:
                continue
            if "最终版" not in filename:
                continue
            if "backup" in filename:
                continue
            docx_files.append(os.path.join(root, filename))
    
    if not docx_files:
        print(f"错误: 在 {source_dir} 及其子目录中未找到包含“技术文件”的docx文件")
        return
        
    print(f"在 {source_dir} 中找到 {len(docx_files)} 个包含“技术文件”的 .docx 文件，即将开始处理...")
    
    # 处理每一个找到的docx文件
    total_files = len(docx_files)
    for i, source_file in enumerate(docx_files, 1):
        print(f"\n[{i}/{total_files}] 正在处理: {os.path.basename(source_file)}")
        basename = os.path.basename(source_file)
        # 跳过临时/锁定文件（如 "~$xxx.docx"、".~-xxx.docx" 等）
        if basename.startswith('~$') or basename.startswith('.~-'):
            print(f"跳过临时文件: {source_file}")
            continue
            
        # 输出目录设置为源文件所在的目录
        output_dir = os.path.dirname(source_file)
        print(f"处理文件目录: {output_dir}")
        try:
            splitter = DocxSplitterV2(source_file, output_dir)
            splitter.run()
        except Exception as e:
            print(f"\n处理文件 {source_file} 时发生未知错误: {e}")
            print("继续处理下一个文件...\n")


if __name__ == "__main__":
    main()
