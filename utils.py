# utils.py
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

def transform_measurement_method(x):
    """將施測方式或儀器代號轉換成 4 位數字字串，再拆分成四個部分"""
    if pd.isnull(x):
        return None
    try:
        s = str(int(x)).zfill(4)
    except Exception:
        s = "0000"
    return {"part1": s[0], "part2": s[1], "part3": s[2], "part4": s[3]}

def set_cell_width(cell, width):
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

def cleanup_temp_files(output_folder, pattern="temp*"):
    """刪除 output_folder 中符合 pattern 的暫存檔案"""
    temp_file_pattern = os.path.join(output_folder, pattern)
    temp_files = glob.glob(temp_file_pattern)
    for temp_file in temp_files:
        try:
            os.remove(temp_file)
            print(f"已刪除暫存檔案: {temp_file}")
        except Exception as e:
            print(f"刪除暫存檔案 {temp_file} 時發生錯誤: {e}")


def overlay_images_to_pdf(original_pdf_path,output_pdf_path):
    """
    利用 ReportLab 與 PyPDF2，從 Tkinter 選取兩張圖片，
    將圖片旋轉後疊加到既有 PDF（固定路徑）的第一頁上，
    並將合併後的結果存到 output_pdf_path 指定的位置。
    """
    # 使用 Tkinter 開啟檔案選取對話框
    root = Tk()
    root.withdraw()  # 隱藏 Tkinter 主視窗

    # 選取第一張圖片
    image_path1 = filedialog.askopenfilename(
        title="請選取監工圖片",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")]
    )
    if not image_path1:
        print("未選取監工圖片，結束。")
        return

    # 選取第二張圖片
    image_path2 = filedialog.askopenfilename(
        title="請選取營業處圖片",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")]
    )
    if not image_path2:
        print("未選取營業處圖片，結束。")
        return

    # --- 步驟 1：利用 ReportLab 生成 overlay PDF (存放在記憶體中) ---
    packet = io.BytesIO()  # 記憶體暫存區
    c = canvas.Canvas(packet, pagesize=A4)

    # 第一張圖片的設定（可依需求調整）
    ratio1 = 0.6
    img_width1 = 177 * ratio1  # 圖片寬度
    img_height1 = 52 * ratio1  # 圖片高度
    angle1 = random.uniform(-3,3)  # 旋轉角度（度）
    center_x1 = 200+random.uniform(-5,5)  # 圖片放置中心 x 座標
    center_y1 = 60+random.uniform(-5,5)  # 圖片放置中心 y 座標

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

    # 第二張圖片的設定（可依需求調整）
    ratio2 = 0.55
    img_width2 = 277 * ratio2
    img_height2 = 181 * ratio2
    angle2 = random.uniform(-3,3)  # 旋轉角度（度）
    center_x2 = 457+random.uniform(-5,5)  # 圖片放置中心 x 座標
    center_y2 = 80+random.uniform(-5,5)  # 圖片放置中心 y 座標

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

    c.save()  # 完成 overlay PDF 的繪製

    # --- 步驟 2：將 overlay 疊加到既有 PDF 上 ---
    # 指定既有的 PDF 路徑（請根據你的需求修改這裡的路徑）
    packet.seek(0)
    overlay_pdf = PdfReader(packet)

    with open(original_pdf_path, "rb") as f_old:
        original_pdf = PdfReader(f_old)
        output = PdfWriter()

        # 這裡假設只在第一頁疊加 overlay，如果需要多頁則可自行調整
        for i, page in enumerate(original_pdf.pages):
            if i == 0:
                page.merge_page(overlay_pdf.pages[0])
            output.add_page(page)

        with open(output_pdf_path, "wb") as f_out:
            output.write(f_out)

    print("PDF 合併完成！輸出檔案：", output_pdf_path)
