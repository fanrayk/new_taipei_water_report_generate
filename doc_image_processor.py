import os
import docx
from docx.shared import Cm
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2pdf import convert
from docx.oxml.ns import qn
from PyPDF2 import PdfMerger

def set_vertical_text_alternative(cell, text):
    """
    將 cell 內容設定為指定文字，
    將每個字拆開後以換行符號串接，達到垂直排列效果，
    並移除段落縮排、水平置中，字型設定為標楷體。
    """
    cell.text = ""
    vertical_text = "\n".join(list(text))
    paragraph = cell.add_paragraph(vertical_text)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.left_indent = 0
    paragraph.paragraph_format.first_line_indent = 0
    run = paragraph.runs[0]
    run.font.name = "標楷體"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")

def get_image_files(plane_folder):
    """
    取得指定資料夾中所有圖片檔案（依副檔名過濾）。
    """
    valid_extensions = [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff"]
    images = [os.path.join(plane_folder, f)
              for f in os.listdir(plane_folder)
              if os.path.splitext(f)[1].lower() in valid_extensions]
    return images

def insert_images_in_template(template_path, image_group, output_file):
    """
    依據模板將圖片與文字插入至表格指定區域：
      - 將第一張圖片插入表格的「第2~5行、第2~3欄」合併儲存格內，並垂直置中。
      - 將第二張圖片插入表格的「第6~9行、第2~3欄」合併儲存格內，並垂直置中。
      - 合併第一欄第2至第9行儲存格，並以垂直排列（每字換行、水平置中）的方式插入文字。
    產生的 Word 文件將儲存至 output_file。
    """
    doc = docx.Document(template_path)
    table = doc.tables[0]

    # 第一張圖片（若有）
    if len(image_group) >= 1:
        merged_cell = table.cell(1, 1).merge(table.cell(4, 2))
        merged_cell.text = ""
        paragraph = merged_cell.paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(image_group[0], width=Cm(15.91))
        merged_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # 第二張圖片（若有）
    if len(image_group) >= 2:
        merged_cell2 = table.cell(5, 1).merge(table.cell(8, 2))
        merged_cell2.text = ""
        paragraph2 = merged_cell2.paragraphs[0]
        run2 = paragraph2.add_run()
        run2.add_picture(image_group[1], width=Cm(15.91))
        merged_cell2.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # 合併第一欄從第2行到第9行，插入垂直排列的文字
    merged_text_cell = table.cell(1, 0).merge(table.cell(8, 0))
    set_vertical_text_alternative(merged_text_cell, "一、竣工平面圖")

    doc.save(output_file)
    print(f"已儲存 Word 文件：{output_file}")

def convert_word_to_pdf(word_path, pdf_path):
    """
    將指定的 Word 文件轉換成 PDF。
    """
    try:
        convert(word_path, pdf_path)
        print(f"已生成 PDF 文件：{pdf_path}")
    except AttributeError as e:
        if "Word.Application.Quit" in str(e):
            print(f"遇到 Word.Application.Quit 錯誤於 {word_path}，忽略並繼續。")
        else:
            raise e

def merge_pdfs(pdf_folder, output_pdf):
    """
    將 pdf_folder 中所有檔名符合 temp_modified_template_group_*.pdf 的 PDF 檔案合併，
    並儲存為 output_pdf。
    """
    pdf_files = [os.path.join(pdf_folder, f)
                 for f in os.listdir(pdf_folder)
                 if f.endswith('.pdf') and f.startswith("temp_modified_template_group_")]
    pdf_files.sort()
    if pdf_files:
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)
        merger.write(output_pdf)
        merger.close()
        print(f"已合併所有 PDF 至：{output_pdf}")

def merge_pdfs_from_list(pdf_list, output_pdf):
    """
    合併由 pdf_list 指定的所有 PDF 檔案，並輸出至 output_pdf。
    
    :param pdf_list: 包含 PDF 檔案路徑的列表，例如 ["1.pdf", "2.pdf", "3.pdf"]
    :param output_pdf: 輸出的 PDF 檔案名稱，例如 "merged.pdf"
    """
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    
    merger.write(output_pdf)
    merger.close()
    print(f"合併完成：{output_pdf}")

def process_documents(main_folder, template_path, output_folder,case_number):
    """
    核心流程：
      1. 從 main_folder 下的「平面圖」資料夾取得所有圖片。
      2. 每兩張圖片分為一組，依據模板產生對應的 Word 文件，
         並轉換為 PDF。
      3. 合併所有 PDF 至一份，並刪除中間產生的 DOCX 檔案。
    """
    plane_folder = os.path.join(main_folder, "平面圖")
    if not os.path.isdir(plane_folder):
        print(f"找不到資料夾：{plane_folder}")
        return

    images = get_image_files(plane_folder)
    print(f"平面圖資料夾中共有 {len(images)} 張圖片。")
    if not images:
        print("沒有找到任何圖片，程式終止。")
        return

    groups = [images[i:i+2] for i in range(0, len(images), 2)]
    print(f"總共分成 {len(groups)} 組。")

    for idx, group in enumerate(groups, start=1):
        word_filename = f"temp_modified_template_group_{idx}.docx"
        word_path = os.path.join(output_folder, word_filename)
        pdf_filename = f"temp_modified_template_group_{idx}.pdf"
        pdf_path = os.path.join(output_folder, pdf_filename)

        insert_images_in_template(template_path, group, word_path)
        convert_word_to_pdf(word_path, pdf_path)

    merged_pdf_path = os.path.join(output_folder, f"temp_{case_number}_平面圖.pdf")
    merge_pdfs(output_folder, merged_pdf_path)
