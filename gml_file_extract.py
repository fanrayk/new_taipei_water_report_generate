import tkinter as tk
from tkinter import filedialog
import xmltodict
import json

# 建立並隱藏主視窗
root = tk.Tk()
root.withdraw()

# 開啟檔案選取對話框，限制檔案類型為 .gml
file_path = filedialog.askopenfilename(
    title="選擇 GML 檔案",
    filetypes=[("GML files", "*.gml")]
)

if file_path:
    with open(file_path, encoding="utf-8") as fd:
        doc = xmltodict.parse(fd.read())

    json_data = json.dumps(doc, indent=4, ensure_ascii=False)
    print(json_data)
else:
    print("未選擇檔案")
