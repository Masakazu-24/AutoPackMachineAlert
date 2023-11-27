import tkinter as tk
from tkinter import ttk
import pandas as pd

def on_name_selected(event):
    selected_name = name_var.get()
    selected_email = name_email_df.loc[name_email_df['Name'] == selected_name, 'mail_head'].values[0]
    email_var.set(selected_email)

# CSVファイルのパス（適切なパスに変更してください）
csv_file_path = 'Name_email.csv'

# Pandasを使用してCSVファイルを読み込む
name_email_df = pd.read_csv(csv_file_path)

root = tk.Tk()
root.title("Name and Email Selection")

# 名前のドロップダウン
name_var = tk.StringVar()
name_label = tk.Label(root, text="Select Name:")
name_label.grid(row=0, column=0, padx=10, pady=10, sticky="E")
name_dropdown = ttk.Combobox(root, textvariable=name_var, values=name_email_df['Name'].tolist())
name_dropdown.grid(row=0, column=1, padx=10, pady=10)
name_dropdown.bind("<<ComboboxSelected>>", on_name_selected)

# 選択された名前に対応するメールアドレスの表示
email_var = tk.StringVar()
email_label = tk.Label(root, text="Selected Email:")
email_label.grid(row=1, column=0, padx=10, pady=10, sticky="E")
email_entry = tk.Entry(root, textvariable=email_var, state="readonly")
email_entry.grid(row=1, column=1, padx=10, pady=10)

# ウィンドウを開始
root.mainloop()
