"""
Excel差异对比工具 - 内部使用版
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.comments import Comment

class ExcelDiffTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel差异对比工具")
        self.root.geometry("600x500")
        
        # 文件选择
        ttk.Label(root, text="旧版文件:").pack(pady=5)
        self.old_file = ttk.Entry(root, width=70)
        self.old_file.pack()
        ttk.Button(root, text="浏览", command=lambda: self.browse(self.old_file)).pack(pady=2)
        
        ttk.Label(root, text="新版文件:").pack(pady=5)
        self.new_file = ttk.Entry(root, width=70)
        self.new_file.pack()
        ttk.Button(root, text="浏览", command=lambda: self.browse(self.new_file)).pack(pady=2)
        
        # 对比按钮
        self.btn = ttk.Button(root, text="开始对比", command=self.start)
        self.btn.pack(pady=10)
        
        # 进度条
        self.progress = ttk.Progressbar(root, mode='indeterminate')
        self.progress.pack(fill='x', padx=20)
        
        # 日志
        self.log = scrolledtext.ScrolledText(root, height=15)
        self.log.pack(fill='both', expand=True, padx=10, pady=10)
        
    def browse(self, entry):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            entry.delete(0, 'end')
            entry.insert(0, f)
    
    def print(self, msg):
        self.log.insert('end', f"[{datetime.now():%H:%M:%S}] {msg}\n")
        self.log.see('end')
        self.root.update()
    
    def start(self):
        if not self.old_file.get() or not self.new_file.get():
            messagebox.showerror("错误", "请选择文件")
            return
        
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not out:
            return
        
        self.btn['state'] = 'disabled'
        self.progress.start()
        threading.Thread(target=self.run, args=(out,), daemon=True).start()
    
    def run(self, out):
        try:
            self.print("加载文件中...")
            old_wb = load_workbook(self.old_file.get())
            new_wb = load_workbook(self.new_file.get())
            result_wb = load_workbook(self.new_file.get())
            
            yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            
            # 对比Sheet
            old_sheets = set(old_wb.sheetnames)
            new_sheets = set(new_wb.sheetnames)
            
            added = new_sheets - old_sheets
            removed = old_sheets - new_sheets
            
            if added:
                self.print(f"新增Sheet: {', '.join(added)}")
            if removed:
                self.print(f"删除Sheet: {', '.join(removed)}")
            
            # 创建汇总Sheet
            if "差异汇总" not in result_wb.sheetnames:
                summary = result_wb.create_sheet("差异汇总", 0)
            else:
                summary = result_wb["差异汇总"]
                summary.delete_rows(1, summary.max_row)
            
            summary['A1'] = "Excel差异对比汇总"
            summary['A1'].font = Font(size=14, bold=True)
            summary['A3'] = f"对比时间: {datetime.now():%Y-%m-%d %H:%M:%S}"
            summary['A4'] = f"旧版: {os.path.basename(self.old_file.get())}"
            summary['A5'] = f"新版: {os.path.basename(self.new_file.get())}"
            
            row_num = 7
            total_diffs = 0
            
            # 对比共有Sheet
            for sheet_name in old_sheets & new_sheets:
                self.print(f"对比: {sheet_name}")
                old_ws = old_wb[sheet_name]
                new_ws = new_wb[sheet_name]
                result_ws = result_wb[sheet_name]
                
                diffs = 0
                max_row = max(old_ws.max_row, new_ws.max_row)
                max_col = max(old_ws.max_column, new_ws.max_column)
                
                for r in range(1, max_row + 1):
                    for c in range(1, max_col + 1):
                        old_cell = old_ws.cell(r, c)
                        new_cell = new_ws.cell(r, c)
                        result_cell = result_ws.cell(r, c)
                        
                        # 快速检查是否相同
                        if old_cell.value == new_cell.value:
                            continue
                        
                        # 有差异
                        diffs += 1
                        result_cell.fill = yellow
                        old_val = str(old_cell.value) if old_cell.value else "(空)"
                        new_val = str(new_cell.value) if new_cell.value else "(空)"
                        result_cell.comment = Comment(
                            f"旧版: {old_val}\n新版: {new_val}", 
                            "ExcelDiff"
                        )
                
                if diffs > 0:
                    self.print(f"  {sheet_name}: {diffs}个差异")
                    summary[f'A{row_num}'] = f"{sheet_name}: {diffs}个差异单元格"
                    row_num += 1
                    total_diffs += diffs
            
            summary[f'A{row_num+1}'] = f"总计: {total_diffs}个差异"
            summary[f'A{row_num+1}'].font = Font(bold=True)
            
            result_wb.save(out)
            self.print(f"完成! 共{total_diffs}个差异, 已保存至: {out}")
            messagebox.showinfo("完成", f"对比完成!\n共{total_diffs}个差异")
            
        except Exception as e:
            self.print(f"错误: {e}")
            messagebox.showerror("错误", str(e))
        finally:
            self.btn['state'] = 'normal'
            self.progress.stop()

if __name__ == "__main__":
    root = tk.Tk()
    ExcelDiffTool(root)
    root.mainloop()
