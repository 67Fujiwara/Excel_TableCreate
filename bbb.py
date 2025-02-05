import tkinter as tk

root = tk.Tk()

# ウィンドウのサイズを設定
root.geometry("800x600")

# Canvasウィジェットを作成
canvas = tk.Canvas(root, width=800, height=600)
canvas.pack(fill="both", expand=True)

# 背景色をグラデーションに設定する関数
def create_gradient(canvas, width, height, color1, color2):
    for i in range(height):
        color = "#%02x%02x%02x" % (
            int(color1[0] + (color2[0] - color1[0]) * i / height),
            int(color1[1] + (color2[1] - color1[1]) * i / height),
            int(color1[2] + (color2[2] - color1[2]) * i / height),
    
        )
        canvas.create_line(0, i, width, i, fill=color)

# グラデーションの開始色と終了色を設定
color1 = (255, 255, 255)  # 白
color2 = (0, 100, 255)    # 青

# グラデーションを作成
create_gradient(canvas, 800, 600, color1, color2)

# ウィジェットをCanvas上に配置する例
button = tk.Button(root, text="クリックしてください")
canvas.create_window(400, 300, window=button)

root.mainloop()
