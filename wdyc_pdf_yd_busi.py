#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word文档拆分工具 - 商务文件版本
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
        self.doc = Document(source_file)

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
        """移除段落前的中文序号 (如 '十一、')"""
        # 只匹配以下形式的序号：
        #   "一、标题" / "二、标题" / "十一、标题" (中文数字 + 顿号)
        #   "1、标题" / "11、标题" (阿拉伯数字 + 顿号)
        # 不匹配：
        #   "2007.2.13" / "71.28%" / "314.43万元" (小数点后的数字)
        
        # 分两种情况：
        # 1. 纯中文数字 + 顿号/括号等
        # 2. 纯阿拉伯数字(不含小数点) + 顿号
        
        # 中文数字序号模式
        chinese_num_pattern = re.compile(r'^[\s\u3000]*[一二三四五六七八九十百千万零〇]+[\s\u3000]*[、）)]')
        # 阿拉伯数字序号模式 (确保后面跟的是顿号，而不是小数点)
        arabic_num_pattern = re.compile(r'^[\s\u3000]*\d+[\s\u3000]*、')
        
        removed_count = 0
        removed_numbering_count = 0
        
        for para in self.iter_all_paragraphs(doc):
            # 1. 移除段落的自动编号格式
            if para._element.pPr is not None:
                numPr = para._element.pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr')
                if numPr is not None:
                    # 移除编号属性
                    para._element.pPr.remove(numPr)
                    removed_numbering_count += 1
                    if para.text.strip():
                        print(f"  移除自动编号: '{para.text[:50]}...'")
            
            # 2. 移除文本中的序号前缀
            if not para.text:
                continue
            
            # 尝试匹配中文数字序号
            match = chinese_num_pattern.match(para.text)
            if not match:
                # 尝试匹配阿拉伯数字序号
                match = arabic_num_pattern.match(para.text)
            
            if not match:
                continue
            
            # 记录原始文本
            original_text = para.text[:50]
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
            removed_count += 1
            print(f"  移除文本序号: '{original_text}...' -> '{para.text[:50]}...'")
        
        if removed_numbering_count > 0:
            print(f"  总共移除了 {removed_numbering_count} 个自动编号")
        if removed_count > 0:
            print(f"  总共移除了 {removed_count} 个文本序号前缀")
    
    def remove_all_numbering_definitions(self, doc):
        """移除文档中的所有编号定义，防止LibreOffice自动添加编号"""
        try:
            # 获取文档的numbering part
            if hasattr(doc, 'part') and hasattr(doc.part, 'numbering_part'):
                numbering_part = doc.part.numbering_part
                if numbering_part is not None:
                    # 清空编号定义
                    numbering_element = numbering_part.element
                    # 移除所有num和abstractNum元素
                    for num in list(numbering_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num')):
                        numbering_element.remove(num)
                    for abstractNum in list(numbering_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNum')):
                        numbering_element.remove(abstractNum)
                    print("  已清除文档中的所有编号定义")
        except Exception as e:
            print(f"  清除编号定义时出错: {e}")
    
    def remove_blank_paragraphs(self, doc):
        """删除文档中的空白段落"""
        try:
            body = doc.element.body
            w_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            
            paragraphs_to_remove = []
            
            for element in list(body):
                if element.tag.endswith('p'):
                    # 1. 检查是否有文本
                    para_text = ''
                    for child in element.iter():
                        if child.tag == f'{w_namespace}t':
                            if child.text:
                                para_text += child.text
                    
                    # 2. 检查是否有图片或对象
                    has_image = False
                    # 检查 drawing (新版Word图片)
                    if element.findall(f'.//{w_namespace}drawing'):
                        has_image = True
                    # 检查 pict (旧版Word图片)
                    elif element.findall(f'.//{w_namespace}pict'):
                        has_image = True
                    # 检查 object (嵌入对象)
                    elif element.findall(f'.//{w_namespace}object'):
                        has_image = True
                    # 检查 vml shape
                    elif element.findall('.//{urn:schemas-microsoft-com:vml}shape'):
                        has_image = True
                    
                    # 如果段落没有文本且没有图片，标记删除
                    if not para_text.strip() and not has_image:
                        paragraphs_to_remove.append(element)
            
            # 删除空白段落
            removed_count = 0
            for element in paragraphs_to_remove:
                body.remove(element)
                removed_count += 1
            
            if removed_count > 0:
                print(f"  已删除 {removed_count} 个空白段落")
            
        except Exception as e:
            print(f"  删除空白段落时出错: {e}")
        
    def get_paragraph_style_level(self, paragraph):
        """获取段落的标题级别"""
        style_name = paragraph.style.name
        if 'Heading' in style_name or '标题' in style_name:
            match = re.search(r'(\d+)', style_name)
            if match:
                return int(match.group(1))
        return None
    
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
            elif start_idx is not None and level:
                para_level = self.get_paragraph_style_level(para)
                if para_level and para_level <= level:
                    end_idx = i
                    break
        
        return start_idx, end_idx
    
    def find_level3_sections_by_parent(self, parent_para_idx):
        """找到某个段落下的所有3级标题章节"""
        sections = []
        
        # 获取父段落的级别
        parent_para = self.doc.paragraphs[parent_para_idx]
        parent_level = self.get_paragraph_style_level(parent_para)
        
        current_section = None
        
        for i in range(parent_para_idx + 1, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            level = self.get_paragraph_style_level(para)
            
            # 如果遇到同级或更高级别的标题，停止
            if level and level <= parent_level:
                if current_section:
                    current_section['end_idx'] = i
                    sections.append(current_section)
                break
            
            # 找到3级标题
            if level == 3:
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
        
        # 添加最后一个section
        if current_section and not current_section['end_idx']:
            current_section['end_idx'] = len(self.doc.paragraphs)
            sections.append(current_section)
        
        return sections
    
    def _cleanup_unused_media(self, docx_path):
        """清理DOCX文件中未使用的媒体文件 (支持 drawingml 和 vml)"""
        try:
            print("  正在清理未使用的媒体文件...")
            
            # XML命名空间
            ns = {
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'v': 'urn:schemas-microsoft-com:vml'
            }
            
            used_rids = set()
            
            # 创建临时目录来解压docx
            with tempfile.TemporaryDirectory() as temp_dir:
                with zipfile.ZipFile(docx_path, 'r') as zin:
                    zin.extractall(temp_dir)
                    
                    # 1. 解析document.xml，找到所有使用的 r:embed 和 r:id ID
                    doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
                    if os.path.exists(doc_xml_path):
                        tree = etree.parse(doc_xml_path)
                        # 查找所有 a:blip 元素下的 r:embed 属性 和 v:imagedata 元素下的 r:id 属性
                        found_ids = tree.xpath("//a:blip/@r:embed | //v:imagedata/@r:id", namespaces=ns)
                        for rid in found_ids:
                            used_rids.add(str(rid))
                
                # 2. 解析rels文件，找到rId对应的图片路径
                rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
                if not os.path.exists(rels_path):
                    print("  未找到rels文件，跳过清理")
                    return
                
                rels_tree = etree.parse(rels_path)
                
                used_media_files = set()
                all_media_rels = {}
                
                # 在rels文件中查找关系
                for rel in rels_tree.iterfind('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rId = rel.get('Id')
                    target = rel.get('Target')
                    
                    # 只关心媒体文件
                    if 'media/' in target:
                        all_media_rels[rId] = target
                        if rId in used_rids:
                            # 规范化路径，例如将 "../media/image1.png" 变为 "word/media/image1.png"
                            used_media_files.add(os.path.normpath(os.path.join('word', target)))
                
                initial_media_count = len(all_media_rels)
                used_media_count = len(used_media_files)
                
                print(f"  找到 {used_media_count} 个正在使用的媒体对象 (共 {initial_media_count} 个)。")

                if used_media_count == initial_media_count:
                    print("  所有媒体文件都在使用中，无需清理。")
                    return
                
                # 3. 创建新的zip文件，只包含使用的文件
                temp_zip_path = docx_path + '.tmp'
                with zipfile.ZipFile(temp_zip_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for dirpath, _, filenames in os.walk(temp_dir):
                        for filename in filenames:
                            file_path = os.path.join(dirpath, filename)
                            archive_name = os.path.relpath(file_path, temp_dir)
                            
                            # 检查是否是未使用的媒体文件
                            is_unused_media = False
                            # 规范化archive_name以进行比较
                            normalized_archive_name = os.path.normpath(archive_name)
                            
                            if normalized_archive_name.startswith('word/media/'):
                                if normalized_archive_name not in used_media_files:
                                    is_unused_media = True
                            
                            if not is_unused_media:
                                zout.write(file_path, archive_name)
                
                # 4. 替换原文件
                shutil.move(temp_zip_path, docx_path)
                print(f"  清理完成：移除了 {initial_media_count - used_media_count} 个未使用的媒体文件。")

        except Exception as e:
            print(f"  清理媒体文件时出错: {e}")
            
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
            last_para_idx = -1  # 记录最后一个段落的索引
            
            for element in list(body):
                # 检查是否是段落
                if element.tag.endswith('p'):
                    should_remove = False
                    last_para_idx = para_count  # 更新最后一个段落的索引
                    
                    # 检查是否在范围外
                    if para_count < start_para_idx or (end_para_idx and para_count >= end_para_idx):
                        should_remove = True
                    else:
                        # 检查段落样式是否包含'toc'（目录样式）
                        if para_count < len(para_styles):
                            style_name = para_styles[para_count]
                            if 'toc' in style_name:
                                should_remove = True
                                print(f"  删除目录样式段落 (idx {para_count}): {style_name}")
                        
                        # 检查是否包含TOC字段指令或超链接字段
                        if not should_remove:
                            for child in element.iter():
                                # 检查instrText元素
                                if child.tag == f'{w_namespace}instrText':
                                    if child.text and ('TOC' in child.text or 'HYPERLINK' in child.text):
                                        should_remove = True
                                        print(f"  删除TOC/HYPERLINK字段段落 (idx {para_count})")
                                        break
                        
                        # 如果段落文本为"目录"，也删除
                        if not should_remove:
                            para_text = ''
                            for child in element.iter():
                                if child.tag == f'{w_namespace}t':
                                    if child.text:
                                        para_text += child.text
                            
                            # 检查是否是目录标题
                            if para_text.strip() in ['目录', 'Table of Contents', 'TOC', '目　录']:
                                should_remove = True
                                print(f"  删除目录标题 (idx {para_count}): {para_text.strip()}")
                            # 检查是否包含目录特征：连续的点号（如 "第一章.......1"）
                            elif '......' in para_text or '..........' in para_text:
                                should_remove = True
                                print(f"  删除目录条目 (idx {para_count}): {para_text[:50]}...")
                            # 检查是否匹配目录条目模式：文字 + 多个点 + 数字
                            elif re.search(r'\.{3,}\s*\d+\s*$', para_text):
                                should_remove = True
                                print(f"  删除目录条目 (idx {para_count}): {para_text[:50]}...")
                    
                    if should_remove:
                        elements_to_remove.append(element)
                    
                    para_count += 1
                    
                # 如果是表格，检查它是否在范围内
                elif element.tag.endswith('tbl'):
                    # 表格紧跟在某个段落之后，使用最后一个段落的索引来判断
                    # 如果最后一个段落在范围内，则表格也在范围内
                    if last_para_idx < start_para_idx or (end_para_idx and last_para_idx >= end_para_idx):
                        elements_to_remove.append(element)
                        print(f"  删除表格 (在段落 {last_para_idx} 之后)")
                    else:
                        print(f"  保留表格 (在段落 {last_para_idx} 之后)")
                
                # 删除SDT元素（通常是目录控件）
                elif element.tag.endswith('sdt'):
                    # SDT元素通常包含目录，直接删除
                    elements_to_remove.append(element)
                    # 获取SDT中的文本以便调试
                    texts = []
                    for t in element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                        if t.text:
                            texts.append(t.text)
                    sdt_text = ''.join(texts)[:100]
                    print(f"  删除SDT元素 (可能是目录): {sdt_text}...")
            
            # 删除不需要的元素
            for element in elements_to_remove:
                body.remove(element)
            
            # 先保存一次，确保文档结构更新
            doc.save(temp_docx)
            
            # 重新加载文档，确保 doc.paragraphs 正确反映当前状态
            doc = Document(temp_docx)
            
            # 移除中文序号前缀
            self.remove_chinese_heading_prefix(doc)
            
            # 移除所有编号定义，防止LibreOffice自动添加编号
            self.remove_all_numbering_definitions(doc)
            
            # 删除空白段落
            self.remove_blank_paragraphs(doc)

            # 再次清理末尾的空白段落，防止出现空白页
            self._remove_trailing_blank_paragraphs(doc)
            
            # 保存最终文档
            doc.save(output_path)
            
            # 清理未使用的媒体文件
            self._cleanup_unused_media(output_path)
    
    def _remove_trailing_blank_paragraphs(self, doc):
        """
        从文档末尾移除连续的空白段落，以防止生成空白页
        """
        w_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        body = doc.element.body
        
        removed_count = 0
        while len(body) > 0:
            last_element = body[-1]
            
            # 只处理段落元素
            if not last_element.tag.endswith('p'):
                break

            # 检查段落是否有文本
            has_text = False
            for t in last_element.iter(f'{w_namespace}t'):
                if t.text and t.text.strip():
                    has_text = True
                    break
            
            # 检查段落是否有图片
            has_drawing = last_element.find(f'.//{w_namespace}drawing') is not None
            has_pict = last_element.find(f'.//{w_namespace}pict') is not None

            # 如果段落既没有文本也没有图片，则移除
            if not has_text and not has_drawing and not has_pict:
                body.remove(last_element)
                removed_count += 1
            else:
                # 一旦遇到有内容的段落，就停止
                break
        
        if removed_count > 0:
            print(f"  已从文档末尾移除 {removed_count} 个空白段落以防止生成空白页。")

    def extract_docx_by_multiple_ranges(self, ranges, output_path):
        """通过复制整个docx文件并只保留指定范围内的段落来提取内容"""
        # ranges 是一个列表，包含 (start_idx, end_idx) 元组
        
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
                    should_remove = True
                    
                    # 检查是否在任何一个范围内
                    for start_idx, end_idx in ranges:
                        if para_count >= start_idx and (end_idx is None or para_count < end_idx):
                            should_remove = False
                            break
                    
                    # 如果在范围内，还需要检查是否是目录或TOC
                    if not should_remove:
                        # 检查段落样式是否包含'toc'（目录样式）
                        if para_count < len(para_styles):
                            style_name = para_styles[para_count]
                            if 'toc' in style_name:
                                should_remove = True
                                print(f"  删除目录样式段落 (idx {para_count}): {style_name}")
                        
                        # 检查是否包含TOC字段指令或超链接字段
                        if not should_remove:
                            for child in element.iter():
                                # 检查instrText元素
                                if child.tag == f'{w_namespace}instrText':
                                    if child.text and ('TOC' in child.text or 'HYPERLINK' in child.text):
                                        should_remove = True
                                        print(f"  删除TOC/HYPERLINK字段段落 (idx {para_count})")
                                        break
                        
                        # 如果段落文本为"目录"，也删除
                        if not should_remove:
                            para_text = ''
                            for child in element.iter():
                                if child.tag == f'{w_namespace}t':
                                    if child.text:
                                        para_text += child.text
                            
                            # 检查是否是目录标题
                            if para_text.strip() in ['目录', 'Table of Contents', 'TOC', '目　录']:
                                should_remove = True
                                print(f"  删除目录标题 (idx {para_count}): {para_text.strip()}")
                            # 检查是否包含目录特征：连续的点号（如 "第一章.......1"）
                            elif '......' in para_text or '..........' in para_text:
                                should_remove = True
                                print(f"  删除目录条目 (idx {para_count}): {para_text[:50]}...")
                            # 检查是否匹配目录条目模式：文字 + 多个点 + 数字
                            elif re.search(r'\.{3,}\s*\d+\s*$', para_text):
                                should_remove = True
                                print(f"  删除目录条目 (idx {para_count}): {para_text[:50]}...")
                    
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
                
                # 处理其他元素 (如 sdt - 可能是目录)
                elif element.tag.endswith('sectPr'):
                    # 保留节属性，以免破坏页面布局
                    pass
                elif element.tag.endswith('sdt'):
                    # SDT元素通常包含目录，直接删除
                    elements_to_remove.append(element)
                    texts = []
                    for t in element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                        if t.text:
                            texts.append(t.text)
                    sdt_text = ''.join(texts)[:100]
                    print(f"  删除SDT元素 (可能是目录): {sdt_text}...")
                else:
                    # 对于其他未知元素，如果当前位置不在范围内，则删除
                    should_remove = True
                    for start_idx, end_idx in ranges:
                        if para_count >= start_idx and (end_idx is None or para_count < end_idx):
                            should_remove = False
                            break
                    
                    if should_remove:
                        elements_to_remove.append(element)
            
            # 删除不需要的元素
            for element in elements_to_remove:
                body.remove(element)
            
            # 先保存一次，确保文档结构更新
            doc.save(temp_docx)
            
            # 重新加载文档，确保 doc.paragraphs 正确反映当前状态
            doc = Document(temp_docx)
            
            # 移除中文序号前缀
            self.remove_chinese_heading_prefix(doc)
            
            # 移除所有编号定义，防止LibreOffice自动添加编号
            self.remove_all_numbering_definitions(doc)
            
            # 删除空白段落
            self.remove_blank_paragraphs(doc)

            # 再次清理末尾的空白段落，防止出现空白页
            self._remove_trailing_blank_paragraphs(doc)
            
            # 保存最终文档
            doc.save(output_path)
            
            # 清理未使用的媒体文件
            self._cleanup_unused_media(output_path)
    
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
            
            # 如果文本长度很短(少于10个字符),认为是空白页
            if len(text) < 10:
                return True
            
            # 检查是否只包含空白字符
            if not text or text.isspace():
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
    
    def extract_section_as_docx(self, section_keyword, output_path):
        """提取特定章节并保存为DOCX文件"""
        print(f"\n提取章节: {section_keyword}")
        
        start_idx = None
        end_idx = None
        
        # 查找起始位置
        for i, para in enumerate(self.doc.paragraphs):
            text = para.text.strip()
            if section_keyword in text:
                # 过滤目录条目
                style_name = para.style.name
                if 'TOC' in style_name or '目录' in style_name:
                    continue
                    
                if '...' in text or re.search(r'\.\s*\d+$', text):
                    continue
                
                start_idx = i
                print(f"  找到章节 '{section_keyword}' at idx {i} (Style: {style_name})")
                break
        
        if start_idx is None:
            print(f"  未找到章节 '{section_keyword}'")
            return False
        
        # 查找结束位置：下一个同级或更高级别的标题
        current_level = self.get_paragraph_style_level(self.doc.paragraphs[start_idx])
        
        for i in range(start_idx + 1, len(self.doc.paragraphs)):
            level = self.get_paragraph_style_level(self.doc.paragraphs[i])
            if level is not None:
                if current_level is not None and level <= current_level:
                    end_idx = i
                    break
                elif current_level is None and level <= 2:
                    end_idx = i
                    break
        
        if end_idx is None:
            end_idx = len(self.doc.paragraphs)
        
        print(f"  范围: {start_idx} - {end_idx}")
        
        # 提取内容到DOCX
        self.extract_docx_by_paragraph_range(start_idx, end_idx, output_path)
        print(f"✓ 已生成: {os.path.basename(output_path)}")
        
        return True
    
    def split_business_sections(self):
        """拆分商务文件特定章节"""
        print("\n1. 拆分商务文件章节...")
        
        # 目标章节关键字
        targets = [
            "法定代表人（单位负责人）授权委托书",
            "商务偏差表",
            "应答人基本情况表",
            "符合采购文件应答人资格要求的证明文件",
            "通用资格",
            "专用资格",
            "诚信评价",
            "综合实力",
            "投标整体响应",
            "重法纪、讲诚信、提质量"
        ]
        
        ranges = []
        
        for target in targets:
            start_idx = None
            end_idx = None
            
            # 查找起始位置
            for i, para in enumerate(self.doc.paragraphs):
                text = para.text.strip()
                if target in text:
                    # 过滤目录条目
                    # 1. 检查样式是否为TOC
                    style_name = para.style.name
                    if 'TOC' in style_name or '目录' in style_name:
                        continue
                        
                    # 2. 检查文本特征：通常目录包含 "......" 或以数字结尾
                    if '...' in text or re.search(r'\.\s*\d+$', text):
                        continue
                        
                    # 3. 优先匹配标题样式 (Heading 1-9)
                    # 如果是普通正文(Normal)，可能是正文中的引用，也可能是标题但没用样式
                    # 我们假设这些主要章节应该是标题，或者至少是独立的行
                    
                    # 找到匹配，记录索引
                    start_idx = i
                    print(f"  锁定 '{target}' at idx {i} (Style: {style_name})")
                    break
            
            if start_idx is not None:
                # 查找结束位置：下一个同级或更高级别的标题
                current_level = self.get_paragraph_style_level(self.doc.paragraphs[start_idx])
                
                # 查找下一个标题作为结束
                for i in range(start_idx + 1, len(self.doc.paragraphs)):
                    level = self.get_paragraph_style_level(self.doc.paragraphs[i])
                    if level is not None:
                        # 如果当前段落有级别，找更高级别或同级
                        if current_level is not None and level <= current_level:
                            end_idx = i
                            break
                        # 如果当前段落没级别，找任何主要标题(1-2级)
                        elif current_level is None and level <= 2: 
                            end_idx = i
                            break
                
                if end_idx is None:
                    end_idx = len(self.doc.paragraphs)
                
                # 针对“符合采购文件应答人资格要求的证明文件”特殊处理：
                # 不包含“（三）查询报告及截图”及之后内容
                if target == "符合采购文件应答人资格要求的证明文件":
                    for i in range(start_idx, end_idx):
                        text = self.doc.paragraphs[i].text.strip()
                        if "（三）查询报告及截图" in text:
                            end_idx = i
                            print("  已截断至“（三）查询报告及截图”之前")
                            break
                
                ranges.append((start_idx, end_idx))
                print(f"  范围: {start_idx} - {end_idx}")
            else:
                print(f"  未找到 '{target}'")
        
        if not ranges:
            print("  未找到任何目标章节")
            return
            
        output_path = os.path.join(self.output_dir, "皖电云采-商务文件汇总.pdf")
        self.save_sections_as_pdf(ranges, output_path)

    def run(self, mode=1):
        """执行拆分任务
        mode=1: 执行老过程（拆分多个章节合并为PDF）
        mode=2: 仅提取"法定代表人（单位负责人）授权委托书"为DOCX
        """
        print("=" * 80)
        if mode == 1:
            print("开始拆分商务文件（完全保留格式）...")
        else:
            print("开始提取授权委托书...")
        print("=" * 80)
        
        if mode == 1:
            self.split_business_sections()
        elif mode == 2:
            # 提取授权委托书
            base_name = os.path.splitext(os.path.basename(self.source_file))[0]
            output_filename = f"授权委托书-{base_name}.docx"
            output_path = os.path.join(self.output_dir, output_filename)
            self.extract_section_as_docx("法定代表人（单位负责人）授权委托书", output_path)
        
        print("\n" + "=" * 80)
        print("所有拆分任务完成！")
        print("=" * 80)


def main():
    # source_dir = "/Users/bisheng/Downloads/lulili2/"
    source_dir = input("请输入源目录路径: ").strip()
    # 移除可能存在的引号（拖拽文件路径到终端时可能会带引号）
    if source_dir.startswith('"') and source_dir.endswith('"'):
        source_dir = source_dir[1:-1]
    if source_dir.startswith("'") and source_dir.endswith("'"):
        source_dir = source_dir[1:-1]
    if not os.path.exists(source_dir):
        print(f"错误: 目录不存在: {source_dir}")
        return
    
    if not os.path.exists('/Applications/LibreOffice.app/Contents/MacOS/soffice'):
        print("错误: 未找到LibreOffice，请先安装LibreOffice")
        print("下载地址: https://www.libreoffice.org/download/download/")
        return
    
    # 显示菜单
    print("\n" + "=" * 80)
    print("请选择处理模式:")
    print("1. 执行老过程（拆分多个章节合并为PDF）")
    print("2. 拆分\"法定代表人（单位负责人）授权委托书\"为单独的DOCX文档")
    print("=" * 80)
    
    mode = None
    while mode not in [1, 2]:
        try:
            mode_input = input("请输入选择 (1 或 2): ").strip()
            mode = int(mode_input)
            if mode not in [1, 2]:
                print("错误: 请输入 1 或 2")
        except ValueError:
            print("错误: 请输入有效的数字")
        except KeyboardInterrupt:
            print("\n\n用户取消操作")
            return
    
    # 递归查找包含"商务文件"或"商务文档"的docx文件
    docx_files = []
    for root, _, files in os.walk(source_dir):
        for filename in files:
            if not filename.lower().endswith(".docx"):
                continue
            # 修改：同时支持"商务文件"和"商务文档"和"商务方案"
            if  "小玲合并版" not in filename:
                continue
            # 跳过已经生成的授权委托书文件 (旧格式)
            if "+授权委托书" in filename:
                continue
            # 跳过已经生成的授权委托书文件 (新格式)
            if filename.startswith("授权委托书-"):
                continue
            full_path = os.path.join(root, filename)
            docx_files.append(full_path)
    
    if not docx_files:
        print(f"错误: 在 {source_dir} 及其子目录中未找到包含\"商务文件\"或\"商务文档\"的docx文件")
        return
    
    print(f"\n找到 {len(docx_files)} 个文件待处理\n")
        
    # 处理每一个找到的docx文件
    for source_file in docx_files:
        basename = os.path.basename(source_file)
        # 跳过临时/锁定文件（如 "~$xxx.docx"、".~-xxx.docx" 等）
        if basename.startswith('~$') or basename.startswith('.~-'):
            print(f"跳过临时文件: {source_file}")
            continue
            
        print(f"\n处理文件: {source_file}")
        
        try:
            output_dir = os.path.dirname(source_file)  # 输出到docx所在目录
            splitter = DocxSplitterV2(source_file, output_dir)
            splitter.run(mode)
        except KeyError as e:
            if "There is no item named" in str(e) and "in the archive" in str(e):
                print(f"✗ 处理失败: 文件已损坏或内部结构不一致。")
                print(f"  错误详情: 文档引用了一个不存在的内部文件（如图片），导致无法读取。")
                print(f"  原始错误: {e}")
                print(f"  已跳过此文件。")
            else:
                print(f"✗ 处理失败: 发生意外的KeyError: {e}")
            continue
        except Exception as e:
            import traceback
            print(f"✗ 处理失败: 发生未知错误。")
            traceback.print_exc()
            continue


if __name__ == "__main__":
    main()
