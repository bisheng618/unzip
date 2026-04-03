import os
from docx import Document
from lxml import etree

def remove_toc(doc_path):
    """
    根据用户提供的结构识别并移除目录。
    支持识别包含 'TOC' 指令的域（instrText）以及标准的 w:sdt 节点。
    """
    try:
        doc = Document(doc_path)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        # 记录所有可能的 TOC 节点及其容器
        # 结果存为 (node_type, node_element) 的列表
        # node_type: 'sdt' 或 'para'
        all_tocs = []
        
        # 1. 查找所有 w:sdt 类型的目录
        body = doc.element.body
        for sdt in body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'):
            sdt_xml = sdt.xml
            if 'Table of Contents' in sdt_xml or 'TOC' in sdt_xml:
                all_tocs.append(('sdt', sdt))
        
        # 2. 查找所有包含 TOC 域代码的段落（非 sdt 包装的情况）
        for para in doc.paragraphs:
            # python-docx 的 xpath 内部已处理 w 命名空间，不需要额外传 namespaces 参数
            if para._element.xpath('.//w:instrText[contains(., "TOC")]'):
                # 检查该段落是否已经在某个已识别的 sdt 内部
                is_inside_already = False
                parent = para._element.getparent()
                while parent is not None:
                    if parent.tag.endswith('sdt'):
                        # 检查此父级 sdt 是否在 all_tocs 中
                        if any(t[1] == parent for t in all_tocs if t[0] == 'sdt'):
                            is_inside_already = True
                            break
                    parent = parent.getparent()
                
                if not is_inside_already:
                    all_tocs.append(('para', para._element))

        if all_tocs:
            count = len(all_tocs)
            print(f"  --> 识别到 {count} 个目录节点")
            if count >= 2:
                # 用户要求：仅删除第二个
                node_type, node_to_delete = all_tocs[1]
                parent = node_to_delete.getparent()
                if parent is not None:
                    parent.remove(node_to_delete)
                    doc.save(doc_path)
                    print(f"  [成功] 已删除第 2 个目录节点 ({node_type})")
                    return True
            else:
                print(f"  [跳过] 仅识别到 {count} 个目录，不符合删除条件。")
                return False
        
        return False

    except Exception as e:
        print(f"  [出错] 处理文件 {doc_path} 时异常: {e}")
        return False

def process_directory(root_dir):
    """
    遍历目录，筛选符合条件的文件并处理
    """
    processed_count = 0
    match_count = 0
    
    # 预处理 root_dir，确保路径格式统一
    root_dir = os.path.abspath(root_dir)
    print(f"\n[开始扫描] 根目录: {root_dir}")
    
    for root, dirs, files in os.walk(root_dir):
        current_dir_name = os.path.basename(root)
        
        # 如果当前路径片段确实叫“上传版”
        if current_dir_name == "上传版":
            # print(f"  正在扫描上传版目录: {root}")
            for file in files:
                # 排除临时文件
                if file.startswith("~$"):
                    continue
                
                # 检查文件名规则：包含“商务文件”和“最终版”，后缀为 .docx
                is_match = "商务文件" in file and "最终版" in file and file.endswith(".docx")
                
                if is_match:
                    match_count += 1
                    file_path = os.path.join(root, file)
                    print(f"\n[匹配成功] 发现目标文件: {file}")
                    print(f"  路径: {file_path}")
                    
                    if remove_toc(file_path):
                        processed_count += 1
                else:
                    # 如果在上传版目录但没匹配，可选打印提示（通常文件很多，这里只打印 docx）
                    if file.endswith(".docx"):
                        pass # 可在此增加调试日志内容
    
    print("\n" + "="*50)
    print(f"处理完成！")
    print(f"符合条件的文件总数: {match_count}")
    print(f"成功移除第二个目录的文件数: {processed_count}")
    print("="*50)

if __name__ == "__main__":
    print("=== Word 文档目录节点移除程序 ===")
    user_input = input("请输入要扫描的根目录路径: ").strip()
    
    if not user_input:
        print("错误：路径不能为空。")
    elif not os.path.exists(user_input):
        print(f"错误：路径不存在 -> {user_input}")
    else:
        process_directory(user_input)
