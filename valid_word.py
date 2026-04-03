import zipfile
import os
import shutil
from xml.etree import ElementTree as ET
from docx import Document
import re

class WordDocValidator:
    """Word文档完整性检查和修复工具"""
    
    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.issues = []
        
    def validate_and_fix(self, output_path=None):
        """验证并尝试修复Word文档"""
        print(f"正在检查文档: {self.docx_path}")
        
        # 1. 基础检查
        if not self._check_basic_structure():
            return False
            
        # 2. 检查图片关系
        self._check_image_relationships()
        
        # 3. 检查XML结构
        self._check_xml_structure()
        
        # 4. 如果有问题，尝试修复
        if self.issues and output_path:
            print(f"\n发现 {len(self.issues)} 个问题，尝试修复...")
            return self._fix_document(output_path)
        
        return len(self.issues) == 0
    
    def _check_basic_structure(self):
        """检查基本的docx结构"""
        try:
            with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
                required_files = [
                    'word/document.xml',
                    '[Content_Types].xml',
                    '_rels/.rels'
                ]
                
                for req_file in required_files:
                    if req_file not in zip_ref.namelist():
                        self.issues.append(f"缺少必需文件: {req_file}")
                        return False
                        
            print("✓ 基本结构完整")
            return True
            
        except zipfile.BadZipFile:
            self.issues.append("文件不是有效的docx格式")
            return False
    
    def _check_image_relationships(self):
        """检查图片关系完整性"""
        try:
            with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
                # 读取关系文件
                if 'word/_rels/document.xml.rels' in zip_ref.namelist():
                    rels_content = zip_ref.read('word/_rels/document.xml.rels')
                    rels_tree = ET.fromstring(rels_content)
                    
                    image_rels = []
                    for rel in rels_tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        rel_type = rel.get('Type', '')
                        if 'image' in rel_type.lower():
                            target = rel.get('Target', '')
                            rel_id = rel.get('Id', '')
                            image_rels.append((rel_id, target))
                    
                    # 检查图片文件是否存在
                    missing_images = []
                    for rel_id, target in image_rels:
                        image_path = f"word/{target}"
                        if image_path not in zip_ref.namelist():
                            missing_images.append((rel_id, target))
                            self.issues.append(f"图片文件缺失: {target} (ID: {rel_id})")
                    
                    if not missing_images:
                        print(f"✓ 找到 {len(image_rels)} 个图片，全部完整")
                    else:
                        print(f"✗ 发现 {len(missing_images)} 个图片引用缺失")
                        
        except Exception as e:
            self.issues.append(f"检查图片关系时出错: {str(e)}")
    
    def _check_xml_structure(self):
        """检查XML结构的有效性"""
        try:
            with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
                # 检查主文档XML
                doc_xml = zip_ref.read('word/document.xml')
                try:
                    tree = ET.fromstring(doc_xml)
                    print("✓ 主文档XML结构有效")
                except ET.ParseError as e:
                    self.issues.append(f"主文档XML解析错误: {str(e)}")
                    
                # 检查样式XML
                if 'word/styles.xml' in zip_ref.namelist():
                    try:
                        styles_xml = zip_ref.read('word/styles.xml')
                        ET.fromstring(styles_xml)
                        print("✓ 样式XML结构有效")
                    except ET.ParseError as e:
                        self.issues.append(f"样式XML解析错误: {str(e)}")
                        
        except Exception as e:
            self.issues.append(f"检查XML结构时出错: {str(e)}")
    
    def _fix_document(self, output_path):
        """尝试修复文档 - 通过重新保存清理结构"""
        try:
            print("\n开始修复文档...")
            
            # 方法1: 使用python-docx重新保存（最简单有效）
            doc = Document(self.docx_path)
            
            # 重新构建文档内容
            new_doc = Document()
            
            # 复制样式和设置
            new_doc.styles._element = doc.styles._element
            new_doc.settings._element = doc.settings._element
            
            # 复制段落（包括图片）
            for paragraph in doc.paragraphs:
                new_para = new_doc.add_paragraph()
                new_para._element = paragraph._element
            
            # 复制表格
            for table in doc.tables:
                new_table = new_doc.add_table(rows=0, cols=len(table.columns))
                new_table._element = table._element
            
            # 保存修复后的文档
            new_doc.save(output_path)
            print(f"✓ 文档已修复并保存到: {output_path}")
            
            # 验证修复后的文档
            validator = WordDocValidator(output_path)
            if validator.validate_and_fix():
                print("✓ 修复成功，文档现在完整")
                return True
            else:
                print("✗ 修复后仍有问题")
                return False
                
        except Exception as e:
            print(f"✗ 修复失败: {str(e)}")
            return False
    
    def print_report(self):
        """打印检查报告"""
        print("\n" + "="*60)
        print("文档检查报告")
        print("="*60)
        
        if not self.issues:
            print("✓ 文档完全正常，可以安全用于程序拼接")
        else:
            print(f"✗ 发现 {len(self.issues)} 个问题：\n")
            for i, issue in enumerate(self.issues, 1):
                print(f"{i}. {issue}")
        
        print("="*60)


def batch_validate_templates(template_dir):
    """批量验证模板文件夹中的所有Word文档"""
    docx_files = [f for f in os.listdir(template_dir) if f.endswith('.docx')]
    
    print(f"找到 {len(docx_files)} 个Word文档\n")
    
    results = {}
    for filename in docx_files:
        filepath = os.path.join(template_dir, filename)
        print(f"\n检查: {filename}")
        print("-" * 60)
        
        validator = WordDocValidator(filepath)
        is_valid = validator.validate_and_fix()
        validator.print_report()
        
        results[filename] = {
            'valid': is_valid,
            'issues': validator.issues
        }
    
    # 总结报告
    print("\n\n" + "="*60)
    print("批量检查总结")
    print("="*60)
    
    valid_count = sum(1 for r in results.values() if r['valid'])
    print(f"完全正常: {valid_count}/{len(results)}")
    print(f"有问题的: {len(results) - valid_count}/{len(results)}")
    
    if len(results) - valid_count > 0:
        print("\n需要修复的文档:")
        for filename, result in results.items():
            if not result['valid']:
                print(f"  - {filename}: {len(result['issues'])} 个问题")


# 使用示例
if __name__ == "__main__":
    # 单个文件检查和修复
    print("示例1: 检查单个文档")
    print("="*60)
    
    # 替换为你的文件路径
    validator = WordDocValidator("/Users/bisheng/Downloads/2121/1766238142_一.docx")
    
    if validator.validate_and_fix(output_path="/Users/bisheng/Downloads/2121/template_fixed.docx"):
        print("\n文档正常或已成功修复")
    else:
        validator.print_report()
    
    print("\n\n示例2: 批量检查文件夹")
    print("="*60)
    
    # 批量检查模板文件夹
    # batch_validate_templates("./templates")