import tkinter as tk
from tkinter import ttk, filedialog
from ttkbootstrap import Style
import pandas as pd
import os
from datetime import datetime

class InventoryManager:
    def __init__(self):
        self.root = tk.Tk()
        self.style = Style(theme='cosmo')
        
        self.root.title("库存管理系统")
        self.root.geometry("800x600")
        
        # 设置变量
        self.inventory_file = tk.StringVar()
        self.operation_file = tk.StringVar()
        self.selected_date = tk.StringVar()
        
        self.setup_ui()
        
        # 绑定文件拖放事件
        self.root.drop_target_register(tk.DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)
        
    def setup_ui(self):
        # 主容器
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text="库存管理系统", 
            font=('Helvetica', 24, 'bold')
        )
        title_label.pack(pady=20)
        
        # 文件选择区域
        files_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        files_frame.pack(fill=tk.X, pady=10)
        
        # 库存文件选择
        inventory_frame = ttk.Frame(files_frame)
        inventory_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(inventory_frame, text="库存文件:").pack(side=tk.LEFT)
        self.inventory_entry = ttk.Entry(
            inventory_frame, 
            textvariable=self.inventory_file,
            width=50
        )
        self.inventory_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            inventory_frame, 
            text="浏览",
            style='info.TButton',
            command=self.select_inventory
        ).pack(side=tk.LEFT)
        
        # 日期选择
        date_frame = ttk.Frame(files_frame)
        date_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(date_frame, text="选择日期:").pack(side=tk.LEFT)
        dates = [str(i) for i in range(1, 32)]
        date_combo = ttk.Combobox(
            date_frame, 
            textvariable=self.selected_date,
            values=dates,
            width=10
        )
        date_combo.pack(side=tk.LEFT, padx=5)
        
        # 出入库文件选择
        operation_frame = ttk.Frame(files_frame)
        operation_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(operation_frame, text="出入库文件:").pack(side=tk.LEFT)
        self.operation_entry = ttk.Entry(
            operation_frame, 
            textvariable=self.operation_file,
            width=50
        )
        self.operation_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            operation_frame, 
            text="浏览",
            style='info.TButton',
            command=self.select_operation
        ).pack(side=tk.LEFT)
        
        # 更新按钮
        ttk.Button(
            main_frame, 
            text="更新库存",
            style='success.TButton',
            command=self.update_inventory,
            width=20
        ).pack(pady=20)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(
            log_frame, 
            height=10,
            yscrollcommand=scrollbar.set,
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
        
    def on_drop(self, event):
        file_path = event.data
        if file_path.endswith(('.xlsx', '.xls')):
            # 根据当前焦点决定更新哪个文件路径
            focused = self.root.focus_get()
            if focused == self.inventory_entry:
                self.inventory_file.set(file_path)
                self.log("已选择库存文件: " + file_path)
            elif focused == self.operation_entry:
                self.operation_file.set(file_path)
                self.log("已选择出入库文件: " + file_path)
    
    def select_inventory(self):
        filename = filedialog.askopenfilename(
            title="选择库存文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.inventory_file.set(filename)
            self.log("已选择库存文件: " + filename)
    
    def select_operation(self):
        filename = filedialog.askopenfilename(
            title="选择出入库文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.operation_file.set(filename)
            self.log("已选择出入库文件: " + filename)
    
    def log(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
    
    def update_inventory(self):
        try:
            if not all([self.inventory_file.get(), self.operation_file.get(), self.selected_date.get()]):
                self.log("错误: 请选择所有必要的文件和日期")
                return
            
            # 读取文件
            df_inventory = pd.read_excel(self.inventory_file.get())
            df_operation = pd.read_excel(self.operation_file.get())
            
            # 判断是出库还是入库
            is_outbound = '出库单号' in df_operation.columns
            operation_type = "出库" if is_outbound else "入库"
            
            # 汇总数量
            if is_outbound:
                df_sum = df_operation.groupby('商品编码')['数量'].sum().reset_index()
            else:
                df_sum = df_operation.groupby('商品编码')['调拨数量'].sum().reset_index()
                df_sum = df_sum.rename(columns={'调拨数量': '数量'})
            
            # 更新库存
            day = int(self.selected_date.get())
            column_name = f"{day}日{'出' if is_outbound else '进'}库"
            updated_count = 0
            not_found_count = 0
            
            for _, row in df_sum.iterrows():
                code = str(row['商品编码'])
                mask = df_inventory['新商品编码'] == code
                if mask.any():
                    df_inventory.loc[mask, column_name] = row['数量']
                    updated_count += 1
                    self.log(f"更新编码 {code} 的{operation_type}数量: {row['数量']}")
                else:
                    not_found_count += 1
                    self.log(f"警告: 未找到编码 {code} 的商品")
            
            # 保存更新后的库存表
            df_inventory.to_excel(self.inventory_file.get(), index=False)
            
            # 显示更新结果
            result = f"更新完成！\n成功更新: {updated_count} 条记录"
            if not_found_count > 0:
                result += f"\n未找到商品: {not_found_count} 条记录"
            
            self.log(result)
            
        except Exception as e:
            self.log(f"错误: {str(e)}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = InventoryManager()
    app.run()
