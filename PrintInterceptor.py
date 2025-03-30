from PIL import Image, ImageTk
import win32print
import win32api
import time
import threading
import os
from tkinter import Tk, Text, Scrollbar, Button, Label, Frame, messagebox

class PrintMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("打印监控程序")
        self.monitoring = False
        self.monitor_thread = None

        # 创建界面组件
        self.label = Label(root, text="打印监控状态: 未启动")
        self.label.pack(pady=10)

        self.start_button = Button(root, text="启动监控", command=self.start_monitoring)
        self.start_button.pack(pady=5)

        self.stop_button = Button(root, text="停止监控", command=self.stop_monitoring, state="disabled")
        self.stop_button.pack(pady=5)

    def start_monitoring(self):
        self.monitoring = True
        self.monitor_thread = threading.Thread(target=self.monitor_print_jobs)
        self.monitor_thread.start()
        self.label.config(text="打印监控状态: 运行中")
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")

    def stop_monitoring(self):
        self.monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join()
        self.label.config(text="打印监控状态: 已停止")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")

    def monitor_print_jobs(self):
        printer_name = win32print.GetDefaultPrinter()
        hprinter = win32print.OpenPrinter(printer_name)
        
        while self.monitoring:
            jobs = win32print.EnumJobs(hprinter, 0, -1, 1)
            if jobs:
                for job in jobs:
                    if job['pDocument']:
                        self.display_file_content(job['pDocument'], job['JobId'])
            time.sleep(1)

    def display_file_content(self, file_name, job_id):
        top = Tk()
        top.title("打印文件内容")
        
        # 检查文件是否存在
        if not os.path.exists(file_name):
            messagebox.showerror("错误", f"文件不存在: {file_name}")
            return
        
        # 判断文件是否为图片
        if file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff')):
            try:
                img = Image.open(file_name)
                img = img.resize((400, 400), Image.ANTIALIAS)  # 调整图片大小
                img_tk = ImageTk.PhotoImage(img)
                
                img_label = Label(top, image=img_tk)
                img_label.image = img_tk  # 保持引用
                img_label.pack()
            except Exception as e:
                messagebox.showerror("错误", f"无法读取图片: {e}")
        else:
            # 文本文件处理逻辑
            text = Text(top, wrap="word")
            scrollbar = Scrollbar(top, command=text.yview)
            text.configure(yscrollcommand=scrollbar.set)
            
            try:
                with open(file_name, 'r', encoding='utf-8') as file:
                    content = file.read()
                    text.insert('1.0', content)
            except Exception as e:
                text.insert('1.0', f"无法读取文件: {e}")
            
            text.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

        # 添加“同意打印”和“驳回打印”按钮
        button_frame = Frame(top)
        button_frame.pack(pady=10)

        def approve_print():
            win32print.SetJob(win32print.OpenPrinter(win32print.GetDefaultPrinter()), job_id, 0, None, win32print.JOB_CONTROL_RELEASE)
            top.destroy()

        def reject_print():
            win32print.SetJob(win32print.OpenPrinter(win32print.GetDefaultPrinter()), job_id, 0, None, win32print.JOB_CONTROL_DELETE)
            top.destroy()

        approve_button = Button(button_frame, text="同意打印", command=approve_print)
        approve_button.pack(side="left", padx=5)

        reject_button = Button(button_frame, text="驳回打印", command=reject_print)
        reject_button.pack(side="right", padx=5)
        
        top.mainloop()

    def monitor_print_jobs(self):
        printer_name = win32print.GetDefaultPrinter()
        hprinter = win32print.OpenPrinter(printer_name)
        
        while self.monitoring:
            jobs = win32print.EnumJobs(hprinter, 0, -1, 1)
            if jobs:
                for job in jobs:
                    if job['pDocument']:
                        self.display_file_content(job['pDocument'], job['JobId'])
            time.sleep(1)

if __name__ == "__main__":
    root = Tk()
    app = PrintMonitorApp(root)
    root.mainloop()