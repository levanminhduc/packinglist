import tkinter as tk
from tkinter import ttk, messagebox
from ui.ui_config import UIConfig


class SettingsDialog:
    
    def __init__(self, parent: tk.Tk, config: UIConfig):
        self.parent = parent
        self.config = config
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Cài Đặt Giao Diện")
        self.dialog.geometry("500x400")
        self.dialog.resizable(False, False)
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._load_current_values()
        
        self._center_window()
    
    def _center_window(self) -> None:
        self.dialog.update_idletasks()
        
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        
        self.dialog.geometry(f'{width}x{height}+{x}+{y}')
    
    def _create_widgets(self) -> None:
        notebook = ttk.Notebook(self.dialog)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        window_frame = ttk.Frame(notebook)
        table_frame = ttk.Frame(notebook)
        
        notebook.add(window_frame, text="Cửa Sổ")
        notebook.add(table_frame, text="Bảng Dữ Liệu")
        
        self._create_window_settings(window_frame)
        self._create_table_settings(table_frame)
        
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            button_frame,
            text="Lưu",
            command=self._save_settings
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            button_frame,
            text="Hủy",
            command=self.dialog.destroy
        ).pack(side=tk.RIGHT)
    
    def _create_window_settings(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Kích Thước Cửa Sổ", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=(10, 5)
        )
        
        ttk.Label(parent, text="Chiều rộng:").grid(row=1, column=0, sticky=tk.W, padx=20, pady=5)
        self.width_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.width_var, width=15).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(parent, text="Chiều cao:").grid(row=2, column=0, sticky=tk.W, padx=20, pady=5)
        self.height_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.height_var, width=15).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(
            row=3, column=0, columnspan=2, sticky=tk.EW, padx=10, pady=10
        )
        
        ttk.Label(parent, text="Vị Trí Cửa Sổ", font=('Arial', 10, 'bold')).grid(
            row=4, column=0, columnspan=2, sticky=tk.W, padx=10, pady=(5, 5)
        )
        
        ttk.Label(parent, text="Vị trí X:").grid(row=5, column=0, sticky=tk.W, padx=20, pady=5)
        self.pos_x_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.pos_x_var, width=15).grid(row=5, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(parent, text="Vị trí Y:").grid(row=6, column=0, sticky=tk.W, padx=20, pady=5)
        self.pos_y_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.pos_y_var, width=15).grid(row=6, column=1, sticky=tk.W, pady=5)
    
    def _create_table_settings(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Kích Thước Cột/Hàng", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=(10, 5)
        )
        
        ttk.Label(parent, text="Độ rộng cột:").grid(row=1, column=0, sticky=tk.W, padx=20, pady=5)
        self.col_width_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.col_width_var, width=15).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(parent, text="Chiều cao hàng:").grid(row=2, column=0, sticky=tk.W, padx=20, pady=5)
        self.row_height_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.row_height_var, width=15).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(
            row=3, column=0, columnspan=2, sticky=tk.EW, padx=10, pady=10
        )
        
        ttk.Label(parent, text="Font Chữ", font=('Arial', 10, 'bold')).grid(
            row=4, column=0, columnspan=2, sticky=tk.W, padx=10, pady=(5, 5)
        )
        
        ttk.Label(parent, text="Kiểu chữ:").grid(row=5, column=0, sticky=tk.W, padx=20, pady=5)
        self.font_family_var = tk.StringVar()
        font_combo = ttk.Combobox(
            parent,
            textvariable=self.font_family_var,
            values=['Arial', 'Courier New', 'Times New Roman', 'Verdana', 'Tahoma'],
            width=13,
            state='readonly'
        )
        font_combo.grid(row=5, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(parent, text="Cỡ chữ:").grid(row=6, column=0, sticky=tk.W, padx=20, pady=5)
        self.font_size_var = tk.StringVar()
        size_combo = ttk.Combobox(
            parent,
            textvariable=self.font_size_var,
            values=['8', '9', '10', '11', '12', '14', '16'],
            width=13,
            state='readonly'
        )
        size_combo.grid(row=6, column=1, sticky=tk.W, pady=5)
        
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(
            row=7, column=0, columnspan=2, sticky=tk.EW, padx=10, pady=10
        )
        
        self.show_grid_var = tk.BooleanVar()
        ttk.Checkbutton(
            parent,
            text="Hiển thị lưới",
            variable=self.show_grid_var
        ).grid(row=8, column=0, columnspan=2, sticky=tk.W, padx=20, pady=5)
    
    def _load_current_values(self) -> None:
        self.width_var.set(str(self.config.get('window.width', 1200)))
        self.height_var.set(str(self.config.get('window.height', 800)))
        self.pos_x_var.set(str(self.config.get('window.position_x', 100)))
        self.pos_y_var.set(str(self.config.get('window.position_y', 50)))
        
        self.col_width_var.set(str(self.config.get('table.column_width', 150)))
        self.row_height_var.set(str(self.config.get('table.row_height', 25)))
        
        self.font_family_var.set(self.config.get('table.font_family', 'Arial'))
        self.font_size_var.set(str(self.config.get('table.font_size', 10)))
        
        self.show_grid_var.set(self.config.get('table.show_grid', True))
    
    def _save_settings(self) -> None:
        try:
            width = int(self.width_var.get())
            height = int(self.height_var.get())
            pos_x = int(self.pos_x_var.get())
            pos_y = int(self.pos_y_var.get())
            
            col_width = int(self.col_width_var.get())
            row_height = int(self.row_height_var.get())
            font_size = int(self.font_size_var.get())
            
            if width < 400 or height < 300:
                messagebox.showerror("Lỗi", "Kích thước cửa sổ quá nhỏ!\nTối thiểu: 400x300")
                return
            
            if col_width < 50 or row_height < 20:
                messagebox.showerror("Lỗi", "Kích thước cột/hàng quá nhỏ!")
                return
            
            self.config.set('window.width', width)
            self.config.set('window.height', height)
            self.config.set('window.position_x', pos_x)
            self.config.set('window.position_y', pos_y)
            
            self.config.set('table.column_width', col_width)
            self.config.set('table.row_height', row_height)
            self.config.set('table.font_family', self.font_family_var.get())
            self.config.set('table.font_size', font_size)
            self.config.set('table.show_grid', self.show_grid_var.get())
            
            messagebox.showinfo("Thành Công", "Đã lưu cài đặt!\nKhởi động lại để áp dụng kích thước cửa sổ.")
            self.dialog.destroy()
            
        except ValueError:
            messagebox.showerror("Lỗi", "Vui lòng nhập số hợp lệ!")

