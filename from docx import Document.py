from docx import Document
from difflib import SequenceMatcher

def get_text_from_docx(file_path):
    """从 Word 文档中提取文本"""
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def find_duplicates(text1, text2, threshold=0.8, min_length=20):
    """比较两段文本，找出重复部分"""
    matcher = SequenceMatcher(None, text1, text2)
    for block in matcher.get_matching_blocks():
        i, j, size = block
        if size >= min_length:
            print(f"重复部分在文档1中：\n{text1[i:i+size]}")
            print(f"重复部分在文档2中：\n{text2[j:j+size]}")
            print("")

def main():
    """主函数"""
    file1_path = r"document1.docx"  # 第一个 Word 文档路径
    file2_path = r"document2.docx"  # 第二个 Word 文档路径

    text1 = get_text_from_docx(file1_path)
    text2 = get_text_from_docx(file2_path)

    find_duplicates(text1, text2)

if __name__ == "__main__":
    main()