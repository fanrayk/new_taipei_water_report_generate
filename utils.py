import os
import random
import glob
import pandas as pd
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
from tkinter import Tk, filedialog
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter

def transform_measurement_method(x) -> dict:
    """
    將施測方式或儀器代號轉換成 4 位數字字串，並拆分成四個部分。
    """
    if pd.isnull(x):
        return None
    try:
        s = str(int(x)).zfill(4)
    except Exception:
        s = "0000"
    return {"part1": s[0], "part2": s[1], "part3": s[2], "part4": s[3]}


def set_cell_width(cell, width: int) -> None:
    """設定 Docx 表格中儲存格的寬度"""
    tc = cell._element
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(width))
    tcW.set(qn("w:type"), "dxa")


def cleanup_temp_files(output_folder: str, pattern: str = "temp*") -> None:
    """刪除 output_folder 中符合 pattern 的暫存檔案"""
    temp_file_pattern = os.path.join(output_folder, pattern)
    temp_files = glob.glob(temp_file_pattern)
    for temp_file in temp_files:
        try:
            os.remove(temp_file)
            print(f"已刪除暫存檔案: {temp_file}")
        except Exception as e:
            print(f"刪除暫存檔案 {temp_file} 時發生錯誤: {e}")


def overlay_images_to_pdf(original_pdf_path: str, output_pdf_path: str) -> None:
    """
    利用 ReportLab 與 PyPDF2，從 Tkinter 選取兩張圖片，
    將旋轉後的圖片疊加到原 PDF 的第一頁上，
    並將結果儲存至 output_pdf_path。
    """
    # 選取圖片
    root = Tk()
    root.withdraw()
    image_path1 = filedialog.askopenfilename(
        title="請選取監工圖片",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")]
    )
    if not image_path1:
        print("未選取監工圖片，結束。")
        return

    image_path2 = filedialog.askopenfilename(
        title="請選取營業處圖片",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")]
    )
    if not image_path2:
        print("未選取營業處圖片，結束。")
        return

    # 生成 overlay PDF 至記憶體
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)

    # 第一張圖片設定
    ratio1 = 0.6
    img_width1 = 177 * ratio1
    img_height1 = 52 * ratio1
    angle1 = random.uniform(-3, 3)
    center_x1 = 200 + random.uniform(-5, 5)
    center_y1 = 60 + random.uniform(-5, 5)
    c.saveState()
    c.translate(center_x1, center_y1)
    c.rotate(angle1)
    c.drawImage(
        image_path1,
        -img_width1 / 2,
        -img_height1 / 2,
        width=img_width1,
        height=img_height1,
        mask="auto",
    )
    c.restoreState()

    # 第二張圖片設定
    ratio2 = 0.55
    img_width2 = 277 * ratio2
    img_height2 = 181 * ratio2
    angle2 = random.uniform(-3, 3)
    center_x2 = 457 + random.uniform(-5, 5)
    center_y2 = 80 + random.uniform(-5, 5)
    c.saveState()
    c.translate(center_x2, center_y2)
    c.rotate(angle2)
    c.drawImage(
        image_path2,
        -img_width2 / 2,
        -img_height2 / 2,
        width=img_width2,
        height=img_height2,
        mask="auto",
    )
    c.restoreState()

    c.save()
    packet.seek(0)
    overlay_pdf = PdfReader(packet)

    with open(original_pdf_path, "rb") as f_old:
        original_pdf = PdfReader(f_old)
        output = PdfWriter()
        for i, page in enumerate(original_pdf.pages):
            if i == 0:
                page.merge_page(overlay_pdf.pages[0])
            output.add_page(page)
        with open(output_pdf_path, "wb") as f_out:
            output.write(f_out)

    print("PDF 合併完成！輸出檔案：", output_pdf_path)


# 以下為從原 main_helpers.py 移入的與檔案處理、圖片產生相關的函式

def generate_dummy_image(filename: str) -> str:
    """
    產生一張全白圖片（500x500），儲存在目前工作目錄，並回傳檔案路徑。
    """
    from PIL import Image
    img = Image.new("RGB", (500, 500), "white")
    dummy_path = os.path.join(os.getcwd(), filename)
    img.save(dummy_path)
    return dummy_path


def process_folder(base_folder: str, folder_name: str, category: str, dummy_prefix: str) -> list:
    """
    處理指定資料夾：若有圖片則回傳圖片列表；否則回傳 dummy 全白圖片。
    若圖片數量為奇數，則補齊 dummy 圖片以湊成偶數。
    """
    folder_path = os.path.join(base_folder, folder_name)
    if os.path.exists(folder_path):
        imgs = []
        for f in os.listdir(folder_path):
            if os.path.splitext(f)[1].lower() in [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff"]:
                imgs.append(os.path.join(folder_path, f))
        if not imgs:
            imgs = [generate_dummy_image(f"blank_{dummy_prefix}.jpg"),
                    generate_dummy_image(f"blank_{dummy_prefix}_2.jpg")]
        elif len(imgs) % 2 == 1:
            imgs.append(generate_dummy_image(f"blank_{dummy_prefix}.jpg"))
    else:
        imgs = [generate_dummy_image(f"blank_{dummy_prefix}.jpg"),
                generate_dummy_image(f"blank_{dummy_prefix}_2.jpg")]
    return [(img, category) for img in imgs]


def process_sorted_folder(base_folder: str, folder_name: str, category: str) -> list:
    """
    處理指定資料夾（依檔名中的數字排序）。
    """
    import re
    folder_path = os.path.join(base_folder, folder_name)
    imgs = []
    if os.path.exists(folder_path):
        files = []
        for f in os.listdir(folder_path):
            if os.path.splitext(f)[1].lower() in [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff"]:
                files.append(os.path.join(folder_path, f))
        imgs = sorted(files, key=lambda x: int(re.search(r"\d+", os.path.basename(x)).group()))
        if len(imgs) % 2 == 1:
            imgs.append(generate_dummy_image(f"blank_{category}.jpg"))
    return [(img, category) for img in imgs]
