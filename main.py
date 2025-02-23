# main.py
import os
from excel_processor import (
    select_folder_and_excel,
    process_excel_pandas,
    process_excel_openpyxl,
    create_output_folder,
)
from doc_generator import (
    generate_records_doc,
    generate_pipeline_doc,
    generate_reserved_doc,
    merge_pdf_files,
)
from utils import cleanup_temp_files, overlay_images_to_pdf


def main():
    # 1. 選取資料夾與 Excel 檔案
    excel_file_path = select_folder_and_excel()

    # 2. 讀取 Excel 資料 (pandas)
    df_renamed = process_excel_pandas(excel_file_path)
    if df_renamed.empty:
        print("Excel 資料讀取失敗，程式結束。")
        exit()
    context_number = df_renamed["case_number"].iloc[0]
    survey_point_count = (
        df_renamed["pipeline_point_count"].iloc[0]
        + df_renamed["manhole_point_count"].iloc[0]
        + df_renamed["facility_point_count"].iloc[0]
    )

    # 3. 建立輸出資料夾
    output_folder = create_output_folder(context_number)

    # # 4. 利用 openpyxl 分離資料
    # simulated_data, reserved_data = process_excel_openpyxl(
    #     excel_file_path, survey_point_count
    # )

    # 5. 產生首頁文件
    records_list = df_renamed.to_dict(orient="records")
    record = records_list[0] if records_list else {}
    records_pdf = generate_records_doc(record, output_folder)

    # # 6. 產生管線文件及（如有）設施物文件
    # pipeline_pdf = generate_pipeline_doc(simulated_data, context_number, output_folder)
    # pdf_list = [records_pdf, pipeline_pdf]
    # if reserved_data:
    #     reserved_pdf = generate_reserved_doc(
    #         reserved_data, context_number, output_folder
    #     )
    #     pdf_list.append(reserved_pdf)
    # merged_pdf_filename = os.path.join(
    #     output_folder, f"temp_{context_number}-附件1-證明資料.pdf"
    # )

    # merge_pdf_files(pdf_list, merged_pdf_filename)

    overlay_images_to_pdf(
        os.path.join(output_folder, f"temp_自主查核表首頁.pdf"),
        os.path.join(output_folder, f"自主查核表首頁.pdf"),
    )

    # 9. 刪除暫存檔案
    cleanup_temp_files(output_folder, "temp*")

    print("========== 全部流程完成 ==========")


if __name__ == "__main__":
    main()
