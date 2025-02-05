import tkinter as tk
from tkinter import ttk
from openpyxl import Workbook

# ウィンドウの作成
root = tk.Tk()
root.title("Excelに表を作成")

# ラベルとエントリの作成
label = tk.Label(root, text="行数:")
label.grid(row=0, column=0)
rows_entry = tk.Entry(root)
rows_entry.grid(row=0, column=1)

label = tk.Label(root, text="列数:")
label.grid(row=1, column=0)
cols_entry = tk.Entry(root)
cols_entry.grid(row=1, column=1)

# ボタンのコールバック関数
def create_table():
    rows = int(rows_entry.get())
    cols = int(cols_entry.get())

    wb = Workbook()
    ws = wb.active

    for i in range(1, rows + 1):
        for j in range(1, cols + 1):
            ws.cell(row=i, column=j, value=f"R{i}C{j}")

    wb.save("table.xlsx")
    result_label.config(text="Excelファイルに表を作成しました")

# ボタンの作成
button = tk.Button(root, text="作成", command=create_table)
button.grid(row=2, columnspan=2)

# 結果表示用ラベル
result_label = tk.Label(root, text="")
result_label.grid(row=3, columnspan=2)

# メインループの開始
root.mainloop()
