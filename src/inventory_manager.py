import tkinter as tk
from tkinter import ttk, filedialog
from ttkbootstrap import Style
import pandas as pd
import os
from datetime import datetime
import traceback
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

class InventoryManager:
    def __init__(self):
        self.root = tk.Tk()
        self.style = Style(theme='cosmo')
        
        self.root.title("库存管理系统")
        self.root.geometry("1000x800")
        
        # 设置变量
        self.inventory_file = tk.StringVar()
        
        self.setup_ui()
        
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
        
        # 工作表选择
        sheet_frame = ttk.Frame(files_frame)
        sheet_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(sheet_frame, text="选择工作表:").pack(side=tk.LEFT)
        self.sheet_combobox = ttk.Combobox(
            sheet_frame,
            width=20,
            state='readonly'
        )
        self.sheet_combobox.pack(side=tk.LEFT, padx=5)
        
        # 出入库文件选择
        operation_frame = ttk.Frame(files_frame)
        operation_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(operation_frame, text="出入库文件:").pack(side=tk.LEFT)
        
        # 使用Listbox显示选择的文件
        self.files_listbox = tk.Listbox(operation_frame, height=5, width=50)
        self.files_listbox.pack(side=tk.LEFT, padx=5)
        
        # 添加文件列表的滚动条
        files_scrollbar = ttk.Scrollbar(operation_frame)
        files_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        
        self.files_listbox.config(yscrollcommand=files_scrollbar.set)
        files_scrollbar.config(command=self.files_listbox.yview)
        
        # 文件操作按钮框架
        file_buttons_frame = ttk.Frame(operation_frame)
        file_buttons_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            file_buttons_frame,
            text="添加文件",
            style='info.TButton',
            command=self.add_operation_files
        ).pack(pady=2)
        
        ttk.Button(
            file_buttons_frame,
            text="清除所选",
            style='danger.TButton',
            command=self.remove_selected_files
        ).pack(pady=2)
        
        ttk.Button(
            file_buttons_frame,
            text="清除全部",
            style='danger.TButton',
            command=self.clear_all_files
        ).pack(pady=2)
        
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
    
    def select_inventory(self):
        filename = filedialog.askopenfilename(
            title="选择库存文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.inventory_file.set(filename)
            self.log("已选择库存文件: " + filename)
            # 读取并更新工作表列表
            try:
                excel_file = pd.ExcelFile(filename)
                self.sheet_combobox['values'] = excel_file.sheet_names
                if len(excel_file.sheet_names) > 0:
                    self.sheet_combobox.set(excel_file.sheet_names[0])
            except Exception as e:
                self.log(f"读取工作表列表时出错: {str(e)}")
    
    def add_operation_files(self):
        filenames = filedialog.askopenfilenames(
            title="选择出入库文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filenames:
            for filename in filenames:
                self.files_listbox.insert(tk.END, filename)
                self.log("已添加文件: " + filename)
    
    def remove_selected_files(self):
        selection = self.files_listbox.curselection()
        for index in reversed(selection):
            self.files_listbox.delete(index)
    
    def clear_all_files(self):
        self.files_listbox.delete(0, tk.END)
    
    def log(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
    
    def update_inventory(self):
        try:
            if not all([self.inventory_file.get(), self.sheet_combobox.get()]):
                self.log("错误: 请选择库存文件和工作表")
                return
                
            if self.files_listbox.size() == 0:
                self.log("错误: 请添加出入库文件")
                return

            # 使用openpyxl读取Excel，保持格式
            wb = load_workbook(self.inventory_file.get())
            selected_sheet = self.sheet_combobox.get()
            ws = wb[selected_sheet]

            # 处理每个出入库文件
            for i in range(self.files_listbox.size()):
                operation_file = self.files_listbox.get(i)
                self.log(f"\n开始处理文件: {os.path.basename(operation_file)}")
                
                try:
                    df_operation = pd.read_excel(operation_file)
                    
                    # 判断是出库还是入库
                    is_outbound = '出库单号' in df_operation.columns
                    operation_type = "出库" if is_outbound else "入库"
                    
                    # 获取日期列名
                    date_column = '出库日期' if is_outbound else '创建日期'
                    
                    if date_column not in df_operation.columns:
                        self.log(f"错误: {operation_type}表中未找到'{date_column}'列")
                        continue
                    
                    # 转换日期列为datetime类型
                    df_operation[date_column] = pd.to_datetime(df_operation[date_column])
                    
                    # 按日期和商品编码分组汇总
                    try:
                        if is_outbound:
                            df_sum = df_operation.groupby([date_column, '商品编码'])['数量'].sum().reset_index()
                        else:
                            df_sum = df_operation.groupby([date_column, '商品编码'])['调拨数量'].sum().reset_index()
                            df_sum = df_sum.rename(columns={'调拨数量': '数量'})
                    except Exception as e:
                        self.log(f"错误: 汇总数量时出错 - {str(e)}")
                        continue

                    # 获取新商品编码列的位置
                    code_column_index = None
                    for idx, cell in enumerate(ws[1], 1):
                        if cell.value == '新商品编码':
                            code_column_index = idx
                            break

                    if code_column_index is None:
                        self.log("错误: 未找到'新商品编码'列")
                        continue

                    # 按日期处理数据
                    for date, group in df_sum.groupby(date_column):
                        day = date.day
                        column_name = f"{day}日{'出' if is_outbound else '进'}库"
                        
                        # 获取列的位置
                        column_index = None
                        for idx, cell in enumerate(ws[1], 1):
                            if cell.value == column_name:
                                column_index = idx
                                break

                        if column_index is None:
                            self.log(f"错误: 未找到列 '{column_name}'")
                            continue

                        updated_count = 0
                        not_found_count = 0

                        # 更新数据
                        for _, row in group.iterrows():
                            code = str(row['商品编码']).strip()
                            quantity = row['数量']
                            
                            found = False
                            for idx, cell in enumerate(ws[get_column_letter(code_column_index)], 1):
                                if str(cell.value).strip() == code:
                                    ws[f"{get_column_letter(column_index)}{idx}"] = quantity
                                    found = True
                                    updated_count += 1
                                    self.log(f"更新编码 {code} 的{date.strftime('%Y-%m-%d')} {operation_type}数量: {quantity}")
                                    break

                            if not found:
                                not_found_count += 1
                                self.log(f"警告: 未找到编码 {code} 的商品")

                        # 显示当前日期的更新结果
                        result = f"\n{date.strftime('%Y-%m-%d')}处理完成！\n成功更新: {updated_count} 条记录"
                        if not_found_count > 0:
                            result += f"\n未找到商品: {not_found_count} 条记录"
                        self.log(result)

                except Exception as e:
                    self.log(f"处理文件 {operation_file} 时出错: {str(e)}")
                    continue

            # 保存更新后的库存表
            try:
                wb.save(self.inventory_file.get())
                self.log("\n已保存更新后的库存表")
            except Exception as e:
                self.log(f"错误: 保存文件时出错 - {str(e)}")
                return

        except Exception as e:
            self.log(f"错误: {str(e)}\n")
            self.log(traceback.format_exc())
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = InventoryManager()
    app.run()
