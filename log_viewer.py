import subprocess
import tkinter as tk
from tkinter import scrolledtext

def run_script():
    # 运行 main_beta.py 并等待其完成
    subprocess.call(["python", "main_beta.py"])

    # 读取日志文件内容
    with open("main_beta.log", "r") as log_file:
        log_content = log_file.read()

    # 显示日志内容
    log_text.insert(tk.END, log_content)
    log_text.see(tk.END)

# 创建主窗口
root = tk.Tk()
root.title("日志查看器")

# 创建一个滚动文本框用于显示日志内容
log_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=30)
log_text.pack(padx=10, pady=10)

# 创建一个按钮用于运行脚本
run_button = tk.Button(root, text="运行脚本", command=run_script)
run_button.pack(pady=10)

# 运行主循环
root.mainloop()
