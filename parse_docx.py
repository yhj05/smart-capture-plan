#2.워드 파일 텍스트 추출
from docx import Document

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    text = "\n".join([para.text for para in doc.paragraphs if para.text.strip() != ""])
    return text
