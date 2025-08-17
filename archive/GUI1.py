import tkinter as tk

# ウィンドウの作成
root = tk.Tk()
root.title("サンプルGUI")
root.geometry("300x200")  # 横300px, 縦200px

# ラベルの追加
label = tk.Label(root, text="こんにちは、世界！")
label.pack(pady=20)

# ボタンの追加
def on_click():
    label.config(text="ボタンが押されました！")

button = tk.Button(root, text="押してみて", command=on_click)
button.pack()

# メインループ開始
root.mainloop()
