# coding='utf-8'
import os
import subprocess
import time
import tkinter as tk
import webbrowser
from tkinter import ttk, messagebox


def open_help_docs():
    """
    打开帮助文档
    """
    help_doc_path = "https://xaviermc.top/?p=43"
    webbrowser.open(help_doc_path)


def check_and_install_pandoc():
    """
    检查并安装Pandoc

    该函数尝试运行`pandoc --version`命令以检测Pandoc是否已安装。
    如果检测到Pandoc，打印消息并返回。
    若未检测到，打印详细安装指导和提示信息。

    注意：在虚拟环境中运行时，可能需要根据实际情况调整Pandoc的安装位置或环境变量设置。
    """
    try:
        subprocess.run(["pandoc", "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
        print("Pandoc已安装")
    except FileNotFoundError:
        print("警告：Pandoc未检测到，请按照以下步骤进行安装：")

        print("\n1. 访问官方发布仓库：")
        print("   https://github.com/jgm/pandoc/releases")
        print("   请在此页面选择适用于您操作系统的最新稳定版本进行下载。")

        print("\n2. 下载完成后，按照官方提供的安装指南或常见安装包的标准流程进行安装。")
        print("   注意确保将Pandoc添加至系统环境变量PATH，以便在任意目录下均可调用。")

        print("\n提示：")
        print("   如您的Python脚本在虚拟环境中运行，由于环境隔离特性，Pandoc的检测结果可能不准确。")
        print("   在这种情况下，请确保在与当前Python脚本所在虚拟环境相匹配或兼容的环境中安装Pandoc，")
        print("   或适当调整虚拟环境配置，使其能够访问宿主机或其他共享环境中的Pandoc安装。")

        print("\n安装完成后，重新运行本程序以确认Pandoc已成功检测到。如有疑问，")
        print("请查阅Pandoc官方文档或联系技术支持获取进一步帮助。")
        print("如您确认Pandoc已安装，但检测结果仍为未安装，可忽略。")


def open_output_folder():
    """
    打开/output文件夹

    该函数检查指定的`output_folder_path`是否存在。
    存在则使用`os.startfile()`打开文件夹；否则打印消息告知用户文件夹不存在。
    """
    output_folder_path = "output"
    if os.path.exists(output_folder_path):
        os.startfile(output_folder_path)
    else:
        print("输出文件夹不存在")


def update_progress(progress, progress_bar):
    """
    更新进度条值并刷新界面

    参数：
    - progress (int): 进度值（0-100）
    - progress_bar (ttk.Progressbar): 要更新的进度条对象
    """
    progress_bar["value"] = progress
    root.update_idletasks()


def simulate_progress(progress_bar):
    """
    模拟进度更新，用于演示进度条功能

    参数：
    - progress_bar (ttk.Progressbar): 要模拟更新的进度条对象
    """
    for _ in range(100):
        update_progress(progress_bar["value"] + 1, progress_bar)
        time.sleep(0.05)


def run_commands(product_id):
    """
    根据商品ID执行一系列命令并更新进度条

    参数：
    - product_id (str): 商品ID
    """
    global progress_bar
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="indeterminate")
    progress_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
    progress_bar.start()

    commands = [
        f"python spider.py -p {product_id}",
        "python clean.py",
        "python segmented.py",
        "python nlp.py",
        'python read_excel.py',
        f"python qianwen.py -p {product_id}"
    ]

    for cmd in commands:
        subprocess.check_call(cmd.split())
        update_progress(10, progress_bar)

    progress_bar.stop()
    messagebox.showinfo("提示", "报告已生成完成！")


def on_click():
    """
    处理"开始分析"按钮点击事件

    获取输入框中的商品ID，如果非空，则调用`run_commands()`；否则显示错误消息。
    """
    product_id = entry.get()
    if product_id:
        run_commands(product_id)
    else:
        messagebox.showerror("错误", "请先输入商品ID")


# 主程序启动

root = tk.Tk()
root.title("JRAS")
root.geometry("600x350")
root.resizable(True, True)
style = ttk.Style()
style.theme_use('clam')

frame = ttk.Frame(root, padding=(20, 10))
frame.pack(pady=10)

ttk.Label(frame, text="请输入京东商品ID号：", font=("Arial", 14)).grid(row=0, column=0, padx=10, pady=10)
entry = ttk.Entry(frame, width=25)
entry.grid(row=0, column=1, padx=10)

start_button = ttk.Button(frame, text="开始分析", command=on_click, width=15)
start_button.grid(row=1, columnspan=2, pady=10)

help_button = ttk.Button(frame, text="帮助文档", command=open_help_docs, width=15)
help_button.grid(row=2, column=0, pady=10)

pandoc_button = ttk.Button(frame, text="检查Pandoc安装", command=check_and_install_pandoc, width=15)
pandoc_button.grid(row=2, column=1, pady=10)

output_button = ttk.Button(frame, text="打开输出文件夹", command=open_output_folder, width=15)
output_button.grid(row=3, columnspan=2, pady=10)

root.mainloop()

if __name__ == '__main__':
    pass