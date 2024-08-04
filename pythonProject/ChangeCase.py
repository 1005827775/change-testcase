import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import ChangeXmind
import threading

def select_xmind_source_file():
    """打开文件对话框，选择xmind文件"""
    try:
        file_path = filedialog.askopenfilename(filetypes=[("xmind files", "*.xmind")])
        if file_path:
            xmind_file_entry.delete(0, tk.END)
            xmind_file_entry.insert(0, file_path)
        else:
            # 提供反馈，告诉用户需要选择一个有效的文件
            tk.messagebox.showinfo("信息", "请选取一个有效的xmind文件.")
    except Exception as e:
        # 异常处理，避免程序崩溃
        tk.messagebox.showerror("错误", f"选择xmind文件时发生错误: {str(e)}")

def choose_excel_save_directory():
    """打开文件夹选择对话框，选择excel文件的保存文件夹"""
    try:
        directory = filedialog.askdirectory()
        if directory:
            excel_path_entry.delete(0, tk.END)
            excel_path_entry.insert(0, directory)
        else:
            tk.messagebox.showinfo("信息", "请指定一个有效的excel保存路径.")
    except Exception as e:
        # 异常处理，避免程序崩溃
        tk.messagebox.showerror("错误", f"选择excel保存路径时发生错误: {str(e)}")

def submit():
    """获取用户输入的数据并执行后续操作"""
    xmind_file = xmind_file_entry.get()
    excel_path = excel_path_entry.get()

    if not xmind_file or not excel_path:
        tk.messagebox.showwarning("警告", "Excel文件或XMind路径未指定.")
        return

    # 使用线程避免阻塞GUI主线程
    threading.Thread(target=perform_submission, args=(xmind_file,excel_path)).start()

def perform_submission(xmind_file, excel_path):
    """执行提交操作"""
    print(f"选择xmind文件: {xmind_file}")
    print(f"选择excel路径: {excel_path}")
    try:
        rs = ChangeXmind.change_to_excel(xmind_file, excel_path)
        tk.messagebox.showinfo("成功", "转换成功:"+rs)
    except Exception as e:
        # 错误处理，记录日志或给用户提示
        tk.messagebox.showerror("错误", f"执行提交操作时发生错误: {str(e)}")

# 创建主窗口
root = tk.Tk()
root.title("测试用例转换")

# 文件选择按钮
xmind_button = tk.Button(root, text="选择Xmind文件", command=select_xmind_source_file)
xmind_button.pack()

# 显示选择的xmind文件路径
xmind_file_entry = tk.Entry(root, width=50)
xmind_file_entry.pack()

# 选择Excel文件保存路径按钮
excel_button = tk.Button(root, text="Excel文件路径", command=choose_excel_save_directory)
excel_button.pack()

excel_path_entry = tk.Entry(root, width=50)
excel_path_entry.pack()

# 提交按钮
submit_button = tk.Button(root, text="生成文件", command=submit)
submit_button.pack()


root.mainloop()
