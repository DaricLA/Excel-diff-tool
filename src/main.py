"""
Excel 差异对比工具（手动打开文件 - 修复跨线程 COM 错误）
- 对比引擎在子线程中重新获取 Excel 实例，避免跨线程调用
- 跳转功能在主线程执行
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import os
import pythoncom
import win32com.client
from win32com.client import constants

# ---------------------------- 辅助函数 ----------------------------
def rgb_to_hex(rgb):
    b = (rgb >> 0) & 0xFF
    g = (rgb >> 8) & 0xFF
    r = (rgb >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"

def normalize_path(path):
    try:
        return os.path.normpath(os.path.realpath(path))
    except:
        return os.path.normpath(path)

def find_workbook(excel_app, target_path):
    """在给定的 Excel 实例中查找指定路径的工作簿"""
    target = normalize_path(target_path)
    for wb in excel_app.Workbooks:
        try:
            if normalize_path(wb.FullName) == target:
                return wb
        except:
            pass
    return None

# ---------------------------- 对比引擎（子线程中独立获取 Excel） ----------------------------
class ExcelComparer:
    def __init__(self, old_path, new_path, log_callback=None, progress_callback=None):
        self.old_path = old_path
        self.new_path = new_path
        self.log = log_callback if log_callback else print
        self.progress = progress_callback if progress_callback else lambda v,s: None
        self.diffs = []
        self.sheet_diffs = []
        self.stats = {
            'total_cells':0, 'diff_cells':0,
            'added_sheets':[], 'removed_sheets':[], 'images_diff':0
        }

    def run(self):
        """在子线程中执行：初始化 COM，获取 Excel 实例，进行对比"""
        pythoncom.CoInitialize()
        excel = None
        try:
            # 在子线程中获取已运行的 Excel 实例
            excel = win32com.client.GetObject(Class="Excel.Application")
            self.log("子线程已连接到 Excel")
            # 输出当前打开的文件
            wb_paths = []
            for wb in excel.Workbooks:
                try:
                    wb_paths.append(wb.FullName)
                except:
                    wb_paths.append("?")
            self.log(f"Excel 中打开的文件: {wb_paths}")

            old_wb = find_workbook(excel, self.old_path)
            new_wb = find_workbook(excel, self.new_path)

            if not old_wb:
                raise Exception(f"未找到旧版文件: {self.old_path}")
            if not new_wb:
                raise Exception(f"未找到新版文件: {self.new_path}")

            self.log("开始对比数据...")
            self.progress(5, "读取数据中...")
            self._compare_sheets(old_wb, new_wb)
            old_sheets = {s.Name: s for s in old_wb.Worksheets}
            new_sheets = {s.Name: s for s in new_wb.Worksheets}
            common = set(old_sheets.keys()) & set(new_sheets.keys())
            total = len(common)
            for idx, name in enumerate(common, 1):
                self.progress(10 + 80*idx//total, f"对比 {name} ({idx}/{total})")
                self.log(f"处理 {name}")
                self._compare_worksheet(old_sheets[name], new_sheets[name], name)

            self.progress(95, "对比完成")
            self.log(f"发现 {self.stats['diff_cells']} 处单元格差异，{len(self.sheet_diffs)} 处 Sheet 差异")
            return True
        except Exception as e:
            self.log(f"对比失败: {e}")
            raise
        finally:
            pythoncom.CoUninitialize()

    # ---------- 比较方法（与之前完全相同）----------
    def _compare_sheets(self, old_wb, new_wb):
        old_names = {s.Name for s in old_wb.Worksheets}
        new_names = {s.Name for s in new_wb.Worksheets}
        self.stats['added_sheets'] = sorted(new_names - old_names)
        self.stats['removed_sheets'] = sorted(old_names - new_names)
        for name in self.stats['added_sheets']:
            self.sheet_diffs.append({'type':'新增Sheet','name':name,'old_name':None,'new_name':name,'desc':f'Sheet "{name}" 只存在于新版'})
        for name in self.stats['removed_sheets']:
            self.sheet_diffs.append({'type':'删除Sheet','name':name,'old_name':name,'new_name':None,'desc':f'Sheet "{name}" 只存在于旧版'})

    def _compare_worksheet(self, old_ws, new_ws, sheet_name):
        old_used = old_ws.UsedRange
        new_used = new_ws.UsedRange
        old_max_row = old_used.Row + old_used.Rows.Count - 1
        old_max_col = old_used.Column + old_used.Columns.Count - 1
        new_max_row = new_used.Row + new_used.Rows.Count - 1
        new_max_col = new_used.Column + new_used.Columns.Count - 1
        max_row = max(old_max_row, new_max_row)
        max_col = max(old_max_col, new_max_col)

        for r in range(1, max_row+1):
            for c in range(1, max_col+1):
                old_cell = old_ws.Cells(r, c) if r <= old_max_row and c <= old_max_col else None
                new_cell = new_ws.Cells(r, c) if r <= new_max_row and c <= new_max_col else None
                addr = old_cell.Address(False, False) if old_cell else new_cell.Address(False, False)
                self.stats['total_cells'] += 1
                diff = self._get_cell_diff(old_cell, new_cell)
                if diff:
                    self.stats['diff_cells'] += 1
                    self.diffs.append({
                        'sheet': sheet_name,
                        'address': addr,
                        'type': diff['type'],
                        'desc': diff['desc']
                    })
        self._compare_merged_cells(old_ws, new_ws, sheet_name)
        self._compare_row_col_dimensions(old_ws, new_ws, sheet_name)
        self._compare_images(old_ws, new_ws, sheet_name)
        self._compare_conditional_formats(old_ws, new_ws, sheet_name)

    def _get_cell_diff(self, old_cell, new_cell):
        if old_cell is None: return {'type':'新增','desc':'旧版无此单元格'}
        if new_cell is None: return {'type':'删除','desc':'新版无此单元格'}
        old_formula = old_cell.Formula
        new_formula = new_cell.Formula
        if old_formula != new_formula:
            return {'type':'公式变化', 'desc': f'公式: {old_formula} → {new_formula}'}
        font_diff = self._cmp_font(old_cell.Font, new_cell.Font)
        fill_diff = self._cmp_interior(old_cell.Interior, new_cell.Interior)
        border_diff = self._cmp_borders(old_cell.Borders, new_cell.Borders)
        align_diff = self._cmp_alignment(old_cell, new_cell)
        num_diff = old_cell.NumberFormat != new_cell.NumberFormat
        merge_diff = old_cell.MergeCells != new_cell.MergeCells
        descs = []
        if font_diff: descs.append(f"字体: {font_diff}")
        if fill_diff: descs.append(f"填充: {fill_diff}")
        if border_diff: descs.append(f"边框: {border_diff}")
        if align_diff: descs.append(f"对齐: {align_diff}")
        if num_diff: descs.append(f"数字格式: {old_cell.NumberFormat} → {new_cell.NumberFormat}")
        if merge_diff: descs.append("合并状态不同")
        if descs: return {'type':'格式变化', 'desc':'; '.join(descs)}
        return None

    def _cmp_font(self, of, nf):
        changes = []
        if of.Name != nf.Name: changes.append(f"字体名: {of.Name}→{nf.Name}")
        if of.Size != nf.Size: changes.append(f"大小: {of.Size}→{nf.Size}")
        if of.Bold != nf.Bold: changes.append(f"粗体: {of.Bold}→{nf.Bold}")
        if of.Italic != nf.Italic: changes.append(f"斜体: {of.Italic}→{nf.Italic}")
        if of.Underline != nf.Underline: changes.append(f"下划线: {of.Underline}→{nf.Underline}")
        if of.Color != nf.Color: changes.append(f"颜色: {rgb_to_hex(of.Color)}→{rgb_to_hex(nf.Color)}")
        return '; '.join(changes) if changes else None

    def _cmp_interior(self, oi, ni):
        if oi.Color != ni.Color or oi.Pattern != ni.Pattern:
            return f"背景: {rgb_to_hex(oi.Color)}/{oi.Pattern} → {rgb_to_hex(ni.Color)}/{ni.Pattern}"
        return None

    def _cmp_borders(self, ob, nb):
        parts = []
        for idx, name in [(7,"左"),(8,"上"),(9,"下"),(10,"右")]:
            o = ob.Item(idx); n = nb.Item(idx)
            if o.LineStyle != n.LineStyle or o.Color != n.Color:
                parts.append(f"{name}: {o.LineStyle}/{rgb_to_hex(o.Color)} → {n.LineStyle}/{rgb_to_hex(n.Color)}")
        return '; '.join(parts) if parts else None

    def _cmp_alignment(self, oc, nc):
        if (oc.HorizontalAlignment != nc.HorizontalAlignment or
            oc.VerticalAlignment != nc.VerticalAlignment or oc.WrapText != nc.WrapText):
            return f"水平: {oc.HorizontalAlignment}→{nc.HorizontalAlignment}, 垂直: {oc.VerticalAlignment}→{nc.VerticalAlignment}, 换行: {oc.WrapText}→{nc.WrapText}"
        return None

    def _compare_merged_cells(self, old_ws, new_ws, sheet_name):
        try:
            old_areas = [a.Address for a in old_ws.UsedRange.MergeAreas]
            new_areas = [a.Address for a in new_ws.UsedRange.MergeAreas]
        except: return
        for addr in set(new_areas)-set(old_areas):
            self.diffs.append({'sheet':sheet_name,'address':addr.split(':')[0],'type':'合并新增','desc':f'新增合并区域 {addr}'})
        for addr in set(old_areas)-set(new_areas):
            self.diffs.append({'sheet':sheet_name,'address':addr.split(':')[0],'type':'合并删除','desc':f'删除合并区域 {addr}'})

    def _compare_row_col_dimensions(self, old_ws, new_ws, sheet_name):
        try:
            old_rows = old_ws.Rows; new_rows = new_ws.Rows
            for r in range(1, old_ws.UsedRange.Rows.Count+1):
                oh = old_rows(r).RowHeight
                nh = new_rows(r).RowHeight if r <= new_ws.UsedRange.Rows.Count else None
                if oh != nh:
                    self.diffs.append({'sheet':sheet_name,'address':f"A{r}",'type':'行高变化','desc':f'行高: {oh} → {nh}'})
        except: pass
        try:
            old_cols = old_ws.Columns; new_cols = new_ws.Columns
            for c in range(1, old_ws.UsedRange.Columns.Count+1):
                ow = old_cols(c).ColumnWidth
                nw = new_cols(c).ColumnWidth if c <= new_ws.UsedRange.Columns.Count else None
                if ow != nw:
                    col_letter = old_ws.Cells(1, c).Address(False, False).replace("1","").replace("$","")
                    self.diffs.append({'sheet':sheet_name,'address':f"{col_letter}1",'type':'列宽变化','desc':f'列宽: {ow} → {nw}'})
        except: pass

    def _compare_images(self, old_ws, new_ws, sheet_name):
        try:
            old_shapes = old_ws.Shapes; new_shapes = new_ws.Shapes
            old_imgs = []; new_imgs = []
            for s in old_shapes:
                if s.Type == 13: old_imgs.append((s.TopLeftCell.Address(False,False), s.Width, s.Height))
            for s in new_shapes:
                if s.Type == 13: new_imgs.append((s.TopLeftCell.Address(False,False), s.Width, s.Height))
            if old_imgs != new_imgs:
                self.stats['images_diff'] += 1
                anchor = old_imgs[0][0] if old_imgs else (new_imgs[0][0] if new_imgs else 'A1')
                self.diffs.append({'sheet':sheet_name,'address':anchor,'type':'图片差异','desc':f'图片数量/尺寸变化'})
        except: pass

    def _compare_conditional_formats(self, old_ws, new_ws, sheet_name):
        try:
            oc = old_ws.UsedRange.FormatConditions.Count
            nc = new_ws.UsedRange.FormatConditions.Count
            if oc != nc:
                self.diffs.append({'sheet':sheet_name,'address':'A1','type':'条件格式数量变化','desc':f'条件格式数量: {oc} → {nc}'})
        except: pass

# ---------------------------- GUI ----------------------------
class DiffViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异对比工具（手动打开文件）")
        self.root.geometry("950x650")
        self.excel_app = None   # 主线程用于跳转的 Excel 实例
        self.old_wb = self.new_wb = None
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()

        top = ttk.Frame(root, padding=10)
        top.pack(fill='x')
        ttk.Label(top, text="旧版文件:").grid(row=0,column=0,sticky='w')
        ttk.Entry(top, textvariable=self.old_path, width=60).grid(row=0,column=1,padx=5)
        ttk.Button(top, text="浏览", command=lambda: self.browse(self.old_path)).grid(row=0,column=2)
        ttk.Label(top, text="新版文件:").grid(row=1,column=0,sticky='w',pady=5)
        ttk.Entry(top, textvariable=self.new_path, width=60).grid(row=1,column=1,padx=5)
        ttk.Button(top, text="浏览", command=lambda: self.browse(self.new_path)).grid(row=1,column=2)

        hint = ttk.Label(top, text="请先用 Excel 打开以上两个文件，然后点击“开始对比”", foreground="blue")
        hint.grid(row=2, column=0, columnspan=3, pady=(5,0))

        btn_frame = ttk.Frame(top)
        btn_frame.grid(row=3,column=0,columnspan=3,pady=5)
        self.start_btn = ttk.Button(btn_frame, text="开始对比", command=self.start_compare)
        self.start_btn.pack(side='left',padx=5)
        self.jump_btn = ttk.Button(btn_frame, text="跳转到选中项", command=self.jump_to_selected, state='disabled')
        self.jump_btn.pack(side='left',padx=5)

        self.progress = ttk.Progressbar(root, mode='determinate')
        self.progress.pack(fill='x',padx=10)

        paned = ttk.PanedWindow(root, orient='horizontal')
        paned.pack(fill='both',expand=True,padx=10,pady=5)
        left = ttk.Frame(paned)
        paned.add(left,weight=3)
        self.tree = ttk.Treeview(left, columns=('address','type'), show='tree headings')
        self.tree.heading('#0',text='Sheet / 差异项'); self.tree.heading('address',text='位置'); self.tree.heading('type',text='类型')
        self.tree.column('#0',width=200); self.tree.column('address',width=80); self.tree.column('type',width=120)
        scroll = ttk.Scrollbar(left, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side='left',fill='both',expand=True); scroll.pack(side='right',fill='y')
        self.tree.bind('<Double-1>', lambda e: self.jump_to_selected())

        right = ttk.Frame(paned)
        paned.add(right,weight=2)
        ttk.Label(right, text="差异详情", font=('微软雅黑',10,'bold')).pack(anchor='w')
        self.detail = tk.Text(right, wrap='word', height=20)
        self.detail.pack(fill='both',expand=True)

        logf = ttk.LabelFrame(root, text="日志", padding=5)
        logf.pack(fill='x',padx=10,pady=(0,10))
        self.log_text = tk.Text(logf, height=6, wrap='word')
        self.log_text.pack(fill='both',expand=True)

        self.diff_items = []

    def browse(self, var):
        p = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if p: var.set(p)

    def log(self, msg):
        self.log_text.insert('end', f"{time.strftime('%H:%M:%S')} {msg}\n")
        self.log_text.see('end')
        self.root.update_idletasks()

    def update_progress(self, val, stat=""):
        self.progress['value'] = val
        if stat: self.log(stat)

    def start_compare(self):
        old = self.old_path.get(); new = self.new_path.get()
        if not old or not new:
            messagebox.showerror("错误","请选择两个文件"); return
        self.start_btn.configure(state='disabled')
        self.tree.delete(*self.tree.get_children())
        self.detail.delete('1.0','end')
        self.diff_items = []

        def worker():
            try:
                comparer = ExcelComparer(old, new, self.log, self.update_progress)
                comparer.run()
                self.result_data = (comparer.diffs, comparer.sheet_diffs, comparer.stats)
                self.root.after(0, self.populate_tree)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("对比失败", str(e)))
            finally:
                self.root.after(0, lambda: self.start_btn.configure(state='normal'))

        threading.Thread(target=worker, daemon=True).start()

    def populate_tree(self):
        diffs, sheet_diffs, stats = self.result_data
        if sheet_diffs:
            sn = self.tree.insert('','end',text='📋 Sheet 结构差异',open=True)
            for sd in sheet_diffs:
                node = self.tree.insert(sn,'end',text=sd['desc'],values=(sd['name'],sd['type']))
                self.diff_items.append((node, {'type':'sheet_struct','data':sd}))
        dmap = {}
        for d in diffs:
            dmap.setdefault(d['sheet'],[]).append(d)
        for sname, items in sorted(dmap.items()):
            pn = self.tree.insert('','end',text=f"📄 {sname}",open=True)
            for d in items:
                node = self.tree.insert(pn,'end',text=d['desc'][:80],values=(d['address'],d['type']))
                self.diff_items.append((node, {'type':'cell','data':d}))
        self.jump_btn.configure(state='normal')
        self.log("树形列表已加载，双击跳转")

    def jump_to_selected(self):
        sel = self.tree.selection()
        if not sel: return
        node = sel[0]
        target = None
        for n, d in self.diff_items:
            if n == node:
                target = d; break
        if not target: return

        if target['type'] == 'cell':
            d = target['data']
            self.detail.delete('1.0','end')
            self.detail.insert('1.0', f"Sheet: {d['sheet']}\n单元格: {d['address']}\n类型: {d['type']}\n描述: {d['desc']}")
            self._navigate(d['sheet'], d['address'])
        elif target['type'] == 'sheet_struct':
            sd = target['data']
            self.detail.delete('1.0','end')
            self.detail.insert('1.0', f"类型: {sd['type']}\n描述: {sd['desc']}")
            if sd.get('new_name'): self._navigate(sd['new_name'], 'A1')
            elif sd.get('old_name'): self._navigate(sd['old_name'], 'A1')

    def _ensure_excel_app(self):
        """在主线程中获取 Excel 实例（用于跳转）"""
        if self.excel_app is None:
            try:
                self.excel_app = win32com.client.GetObject(Class="Excel.Application")
            except:
                messagebox.showerror("错误", "未检测到正在运行的 Excel，请先打开 Excel 文件。")
                return False
        return True

    def _navigate(self, sheet_name, address):
        if not self._ensure_excel_app():
            return
        try:
            old_wb = find_workbook(self.excel_app, self.old_path.get())
            new_wb = find_workbook(self.excel_app, self.new_path.get())
            for wb, desc in [(old_wb, "旧版"), (new_wb, "新版")]:
                if wb is None:
                    self.log(f"{desc} 文件未打开，无法跳转")
                    continue
                try:
                    ws = wb.Worksheets(sheet_name)
                    ws.Activate()
                    ws.Range(address).Select()
                    self.log(f"{desc} 已跳转到 {sheet_name}!{address}")
                except Exception as e:
                    self.log(f"{desc} 跳转失败: {e}")
        except Exception as e:
            messagebox.showerror("跳转错误", f"无法操作 Excel:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DiffViewer(root)
    root.mainloop()
