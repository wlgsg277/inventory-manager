import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry

class InventoryManager:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("库存管理系统")
        self.window.geometry("800x600")
        
        # 设置变量
        self.inventory_file = tk.StringVar()
        self.selected_file = tk.StringVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 库存文件选择区域
        inventory_frame = ttk.LabelFrame(main_frame, text="国内电商库存文件", padding="5")
        inventory_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(inventory_frame, textvariable=self.inventory_file, width=60).pack(side=tk.LEFT, padx=5)
        ttk.Button(inventory_frame, text="选择库存文件", command=self.select_inventory).pack(side=tk.LEFT, padx=5)
        
        # 出入库文件选择区域
        operation_frame = ttk.LabelFrame(main_frame, text="出入库文件", padding="5")
        operation_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(operation_frame, textvariable=self.selected_file, width=60).pack(side=tk.LEFT, padx=5)
        ttk.Button(operation_frame, text="选择文件", command=self.select_file).pack(side=tk.LEFT, padx=5)
        
        # 更新按钮
        ttk.Button(main_frame, text="更新库存", command=self.update_inventory).pack(pady=10)
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, height=20, yscrollcommand=scrollbar.set)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
        
    def select_inventory(self):
        filename = filedialog.askopenfilename(
            title="选择库存文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.inventory_file.set(filename)
            self.log("已选择库存文件: " + filename)
    
    def select_file(self):
        filename = filedialog.askopenfilename(
            title="选择出入库文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.selected_file.set(filename)
            self.log("已选择文件: " + filename)
    
    def log(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
    
    def update_inventory(self):
        try:
            if not self.inventory_file.get() or not self.selected_file.get():
                messagebox.showerror("错误", "请选择所有必要的文件")
                return
            
            # 读取文件
            df_inventory = pd.read_excel(self.inventory_file.get())
            df_operation = pd.read_excel(self.selected_file.get())
            
            # 获取日期
            date_str = df_operation.iloc[0]['创建日期']
            day = pd.to_datetime(date_str).day
            
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
            result_message = f"更新完成！\n"
            result_message += f"成功更新: {updated_count} 条记录\n"
            if not_found_count > 0:
                result_message += f"未找到商品: {not_found_count} 条记录"
            
            self.log(f"完成更新 - {result_message}")
            messagebox.showinfo("成功", result_message)
            
        except Exception as e:
            error_message = f"错误: {str(e)}"
            self.log(error_message)
            messagebox.showerror("错误", f"更新失败:\n{error_message}")
    
    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = InventoryManager()
    app.run()
