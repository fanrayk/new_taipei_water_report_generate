import os
from docx2pdf import convert as docx2pdf_convert
from PyPDF2 import PdfMerger

from excel_processor import select_folder_and_excel, process_excel_pandas, create_output_folder
from doc_generator import generate_records_doc
from doc_image_processor import process_documents, merge_pdfs_from_list, insert_images_into_9x3_template_left_to_right
from utils import cleanup_temp_files, overlay_images_to_pdf, process_folder, process_sorted_folder
from config import TEMPLATE_TABLE

def main():
    # 1. 選取 Excel 檔案所在資料夾與檔案
    excel_file_path = select_folder_and_excel()

    # 2. 讀取 Excel 資料
    df_renamed = process_excel_pandas(excel_file_path)
    if df_renamed.empty:
        print("Excel 資料讀取失敗，程式結束。")
        exit()
    context_number = df_renamed["case_number"].iloc[0]

    # 3. 建立輸出資料夾
    output_folder = create_output_folder(context_number)

    # 4. 產生首頁文件與疊加圖片
    records_pdf = generate_records_doc(df_renamed.to_dict(orient="records")[0], output_folder)
    overlay_images_to_pdf(
        os.path.join(output_folder, "temp_自主查核表首頁.pdf"),
        os.path.join(output_folder, f"temp_{context_number}_自主查核表首頁.pdf")
    )

    # 5. 處理平面圖文件並合併 PDF
    main_folder = os.path.dirname(excel_file_path)
    process_documents(main_folder, TEMPLATE_TABLE, output_folder, context_number)

    # 6. 處理各類照片
    base_folder = os.path.dirname(excel_file_path)
    images = []
    images.extend(process_folder(base_folder, "埋深照", "埋深照", "0"))
    images.extend(process_folder(base_folder, "銑鋪照", "銑鋪照", "1"))
    images.extend(process_sorted_folder(base_folder, "測量照", "測量照"))
    images.extend(process_sorted_folder(base_folder, "讀數照", "讀數照"))

    if not images:
        print("找不到任何圖片，程式結束。")
        return

    template_path = os.path.join("template", "自主查核表_表格模板.docx")
    output_prefix = os.path.join(output_folder, str(context_number))
    word_files = insert_images_into_9x3_template_left_to_right(template_path, images, output_prefix)

    # 7. 轉換所有 Word 檔為 PDF
    pdf_files = []
    for word_file in word_files:
        pdf_file = word_file.replace(".docx", ".pdf")
        docx2pdf_convert(word_file, pdf_file)
        print(f"已轉換為 PDF：{pdf_file}")
        pdf_files.append(pdf_file)

    # 8. 合併其他照片 PDF
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    merged_pdf_path = os.path.join(output_folder, f"temp_{context_number}_其他照片.pdf")
    merger.write(merged_pdf_path)
    merger.close()
    print(f"已合併 PDF：{merged_pdf_path}")

    # 9. 合併最終 PDF
    merge_pdfs_from_list(
        [
            os.path.join(output_folder, f"temp_{context_number}_自主查核表首頁.pdf"),
            os.path.join(output_folder, f"temp_{context_number}_平面圖.pdf"),
            os.path.join(output_folder, f"temp_{context_number}_其他照片.pdf"),
        ],
        os.path.join(output_folder, f"{context_number}_自主查核表.pdf")
    )

    # 10. 刪除暫存檔案
    cleanup_temp_files(output_folder, "temp*")
    cleanup_temp_files(os.getcwd(), "blank*")

    print("========== 全部流程完成 ==========")

if __name__ == "__main__":
    main()
