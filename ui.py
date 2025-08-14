import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import van  # Import module chính
import os

class ImageExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Xuất Ảnh Thẻ từ Excel")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Biến lưu trữ đường dẫn
        self.excel_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.scale_factor = tk.DoubleVar(value=3.0)
        self.wait_time = tk.DoubleVar(value=0.5)
        
        # Tạo giao diện
        self.create_widgets()
        
        # Thiết lập giá trị mặc định
        self.output_folder.set("ANHTHE")
        
    def create_widgets(self):
        # Tạo notebook (tab)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab chính
        main_frame = ttk.Frame(notebook)
        notebook.add(main_frame, text='Cài đặt Xuất Ảnh')
        
        # Tab log
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text='Nhật ký Hoạt động')
        
        # Tạo giao diện cho tab chính
        self.create_main_tab(main_frame)
        
        # Tạo giao diện cho tab log
        self.create_log_tab(log_frame)
    
    def create_main_tab(self, parent):
        # Frame chứa các control
        control_frame = ttk.LabelFrame(parent, text="Thiết lập Xuất Ảnh")
        control_frame.pack(fill='x', padx=10, pady=10)
        
        # File Excel
        ttk.Label(control_frame, text="File Excel:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        excel_entry = ttk.Entry(control_frame, textvariable=self.excel_file, width=50)
        excel_entry.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        ttk.Button(control_frame, text="Chọn...", command=self.browse_excel).grid(row=0, column=2, padx=5, pady=5)
        
        # Thư mục đầu ra
        ttk.Label(control_frame, text="Thư mục Ảnh:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        output_entry = ttk.Entry(control_frame, textvariable=self.output_folder, width=50)
        output_entry.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        ttk.Button(control_frame, text="Chọn...", command=self.browse_output).grid(row=1, column=2, padx=5, pady=5)
        
        # Hệ số phóng to
        ttk.Label(control_frame, text="Hệ số Phóng to:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        scale_spin = ttk.Spinbox(control_frame, textvariable=self.scale_factor, from_=1.0, to=10.0, increment=0.5, width=10)
        scale_spin.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        
        # Thời gian chờ
        ttk.Label(control_frame, text="Thời gian Chờ (giây):").grid(row=3, column=0, sticky='w', padx=5, pady=5)
        wait_spin = ttk.Spinbox(control_frame, textvariable=self.wait_time, from_=0.1, to=5.0, increment=0.1, width=10)
        wait_spin.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        # Nút bắt đầu
        start_frame = ttk.Frame(parent)
        start_frame.pack(fill='x', padx=10, pady=10)
        
        self.start_button = ttk.Button(start_frame, text="Bắt Đầu Xuất Ảnh", command=self.start_export)
        self.start_button.pack(pady=10)
        
        # Trạng thái
        self.status_var = tk.StringVar(value="Sẵn sàng")
        status_bar = ttk.Label(parent, textvariable=self.status_var, relief='sunken', anchor='w')
        status_bar.pack(side='bottom', fill='x', padx=10, pady=10)
    
    def create_log_tab(self, parent):
        # Tạo textbox cho log
        self.log_text = scrolledtext.ScrolledText(parent, wrap=tk.WORD)
        self.log_text.pack(fill='both', expand=True, padx=10, pady=10)
        self.log_text.config(state='disabled')
        
        # Nút xóa log
        clear_button = ttk.Button(parent, text="Xóa Nhật ký", command=self.clear_log)
        clear_button.pack(side='bottom', pady=10)
    
    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file.set(file_path)
    
    def browse_output(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder.set(folder_path)
    
    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
    
    def log_message(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
    
    def update_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def start_export(self):
        # Kiểm tra file Excel
        if not self.excel_file.get():
            messagebox.showerror("Lỗi", "Vui lòng chọn file Excel!")
            return
        
        # Tắt nút bắt đầu trong khi xử lý
        self.start_button.config(state='disabled')
        self.log_message("=" * 50)
        self.log_message("BẮT ĐẦU QUÁ TRÌNH XUẤT ẢNH")
        self.log_message("=" * 50)
        
        # Lấy tham số
        params = {
            'excel_file_path': self.excel_file.get(),
            'output_folder': self.output_folder.get(),
            'scale_factor': self.scale_factor.get(),
            'wait_time': self.wait_time.get()
        }
        
        # Chạy trong luồng riêng để không làm đơ giao diện
        thread = threading.Thread(target=self.run_export, args=(params,))
        thread.daemon = True
        thread.start()
    
    def run_export(self, params):
        try:
            # Tạo hàm callback để cập nhật UI
            def log_callback(message):
                self.log_message(message)
                self.update_status(message)
            
            # Chạy hàm xuất ảnh từ module van
            van.export_images(
                excel_file_path=params['excel_file_path'],
                output_folder=params['output_folder'],
                scale_factor=params['scale_factor'],
                wait_time=params['wait_time'],
                log_callback=log_callback
            )
            
            self.log_message("\n" + "=" * 50)
            self.log_message("HOÀN TẤT XUẤT ẢNH!")
            self.log_message("=" * 50)
            
            # Mở thư mục kết quả
            if os.path.exists(params['output_folder']):
                os.startfile(params['output_folder'])
            
        except Exception as e:
            self.log_message(f"\n❌ LỖI: {str(e)}")
        finally:
            # Kích hoạt lại nút bắt đầu
            self.start_button.config(state='normal')
            self.update_status("Hoàn tất - Sẵn sàng cho lần xuất tiếp theo")

def main():
    root = tk.Tk()
    app = ImageExportApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()