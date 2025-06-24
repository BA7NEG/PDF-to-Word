import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pdf2docx import Converter
import time

class PDFToWordConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF转Word工具 v1.0")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # 设置窗口图标（在Windows环境中使用）
        # self.root.iconbitmap("icon.ico")  # 在Windows环境中取消注释
        
        # 设置样式
        self.style = ttk.Style()
        self.style.configure("TButton", font=("微软雅黑", 10))
        self.style.configure("TLabel", font=("微软雅黑", 10))
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件选择部分
        self.file_frame = ttk.LabelFrame(self.main_frame, text="文件选择", padding="10")
        self.file_frame.pack(fill=tk.X, pady=10)
        
        self.file_path = tk.StringVar()
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_button = ttk.Button(self.file_frame, text="浏览", command=self.browse_file)
        self.browse_button.pack(side=tk.RIGHT, padx=5)
        
        # 输出目录选择部分
        self.output_frame = ttk.LabelFrame(self.main_frame, text="输出目录", padding="10")
        self.output_frame.pack(fill=tk.X, pady=10)
        
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_path, width=50)
        self.output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.output_button = ttk.Button(self.output_frame, text="浏览", command=self.browse_output)
        self.output_button.pack(side=tk.RIGHT, padx=5)
        
        # 进度条部分
        self.progress_frame = ttk.LabelFrame(self.main_frame, text="转换进度", padding="10")
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=100, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        self.status_label = ttk.Label(self.progress_frame, textvariable=self.status_var)
        self.status_label.pack(pady=5)
        
        # 转换按钮
        self.convert_button = ttk.Button(self.main_frame, text="开始转换", command=self.start_conversion)
        self.convert_button.pack(pady=10)
        
        # 结果报告部分
        self.report_frame = ttk.LabelFrame(self.main_frame, text="转换结果", padding="10")
        self.report_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.report_text = tk.Text(self.report_frame, height=5, wrap=tk.WORD)
        self.report_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        self.scrollbar = ttk.Scrollbar(self.report_text, command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 添加制作人信息
        self.author_label = ttk.Label(self.main_frame, text="制作人: BA7NEG | 版本: v1.0", font=("微软雅黑", 8))
        self.author_label.pack(side=tk.BOTTOM, pady=5)
        
        # 转换线程
        self.conversion_thread = None
        
    def browse_file(self):
        """浏览并选择PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)
            # 默认设置输出目录为PDF文件所在目录
            self.output_path.set(os.path.dirname(file_path))
    
    def browse_output(self):
        """浏览并选择输出目录"""
        output_dir = filedialog.askdirectory(title="选择输出目录")
        if output_dir:
            self.output_path.set(output_dir)
    
    def start_conversion(self):
        """开始转换过程"""
        pdf_path = self.file_path.get()
        output_dir = self.output_path.get()
        
        # 验证输入
        if not pdf_path:
            messagebox.showerror("错误", "请选择PDF文件")
            return
        
        if not output_dir:
            messagebox.showerror("错误", "请选择输出目录")
            return
        
        if not os.path.exists(pdf_path):
            messagebox.showerror("错误", "PDF文件不存在")
            return
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录: {str(e)}")
                return
        
        # 禁用按钮，防止重复点击
        self.convert_button.configure(state=tk.DISABLED)
        self.browse_button.configure(state=tk.DISABLED)
        self.output_button.configure(state=tk.DISABLED)
        
        # 清空报告
        self.report_text.delete(1.0, tk.END)
        
        # 重置进度条并开始动画
        self.progress.start(10)
        self.status_var.set("正在转换...")
        
        # 在新线程中执行转换，避免界面卡死
        self.conversion_thread = threading.Thread(target=self.convert_pdf_to_word, args=(pdf_path, output_dir))
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
    
    def convert_pdf_to_word(self, pdf_path, output_dir):
        """在后台线程中执行PDF到Word的转换"""
        try:
            # 获取文件名（不含扩展名）
            file_name = os.path.basename(pdf_path)
            file_name_without_ext = os.path.splitext(file_name)[0]
            
            # 构建输出文件路径
            docx_path = os.path.join(output_dir, f"{file_name_without_ext}.docx")
            
            # 添加转换开始信息
            self.update_report(f"开始转换: {file_name}")
            self.update_report(f"输出文件: {docx_path}")
            
            # 记录开始时间
            start_time = time.time()
            
            # 创建转换器并执行转换
            # 不再尝试获取页数，直接进行转换
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
            
            # 计算耗时
            elapsed_time = time.time() - start_time
            
            # 更新UI
            self.root.after(0, lambda: self.update_status("转换完成!"))
            self.root.after(0, lambda: self.update_report(f"转换完成! 耗时: {elapsed_time:.2f} 秒"))
            self.root.after(0, lambda: self.conversion_completed())
            
        except Exception as e:
            # 更新UI显示错误
            self.root.after(0, lambda: self.update_status(f"转换失败: {str(e)}"))
            self.root.after(0, lambda: self.update_report(f"错误: {str(e)}"))
            self.root.after(0, lambda: self.conversion_completed(success=False))
    
    def update_status(self, message):
        """更新状态标签"""
        self.status_var.set(message)
    
    def update_report(self, message):
        """更新结果报告"""
        self.report_text.insert(tk.END, message + "\n")
        self.report_text.see(tk.END)  # 自动滚动到最新内容
    
    def conversion_completed(self, success=True):
        """转换完成后的处理"""
        # 停止进度条动画
        self.progress.stop()
        
        # 重新启用按钮
        self.convert_button.configure(state=tk.NORMAL)
        self.browse_button.configure(state=tk.NORMAL)
        self.output_button.configure(state=tk.NORMAL)
        
        if success:
            # 设置进度条为满
            self.progress['value'] = 100
            
            # 显示成功消息
            messagebox.showinfo("成功", "PDF转换为Word文档成功!")
            
            # 询问是否打开输出目录
            if messagebox.askyesno("打开文件夹", "是否打开输出目录?"):
                output_dir = self.output_path.get()
                # 在Windows系统中使用以下命令打开文件夹
                # 在实际Windows环境中会执行，在其他环境中可能会失败
                try:
                    os.startfile(output_dir)
                except:
                    self.update_report("无法自动打开输出目录，请手动查看。")
        else:
            # 重置进度条
            self.progress['value'] = 0

def main():
    root = tk.Tk()
    app = PDFToWordConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
