import os
from docxtpl import DocxTemplate
from docx2pdf import convert
from config import TEMPLATE_MAIN

def generate_records_doc(record, output_folder):
    """
    使用模板產生 Word 文件，並轉換為 PDF。
    """
    doc = DocxTemplate(TEMPLATE_MAIN)
    doc.render(record)
    docx_path = os.path.join(output_folder, "temp_自主查核表首頁.docx")
    pdf_path = os.path.join(output_folder, "temp_自主查核表首頁.pdf")

    doc.save(docx_path)
    convert(docx_path, pdf_path)

    print(f"Records PDF 已產生：{pdf_path}")
    return pdf_path
