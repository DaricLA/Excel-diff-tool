"""
Excel 差异对比工具（win32com 原生版）
- 使用 win32com 读取新旧 Excel，完全不修改文件
- 树形列表展示差异，按 Sheet 分组
- 双击跳转：两个 Excel 窗口同步选中差异单元格（或图片锚点）
- 支持 Excel 2010+，需要安装 Microsoft Excel
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import time
import pythoncom
import win32com.client
from win32com.client import constants

# ---------------------------- 辅助函数 ----------------------------
def rgb_to_hex(rgb):
    """将 BGR 整数转换为 #RRGGBB 字符串（用于显示）"""
    b = (rgb >> 0) & 0xFF
    g = (rgb >> 8) & 0xFF
    r = (rgb >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"

# ---------------------------- Excel 读取与对比引擎 ----------------------------
class ExcelComparer:
    def __init__(self, old_path, new_path, log_callback=None, progress_callback=None):
        self.old_path = old_path
        self.new_path = new_path
        self.log = log_callback if log_callback else print
        self.progress = progress_callback if progress_callback else lambda v, s: None
        self.diffs = []           # 所有差异项，格式：{sheet, address, type, desc, old_ws_name, new_ws_name}
        self.sheet_diffs = []     # Sheet 结构差异：{type, name, desc}
        self.stats = {
            'total_cells': 0,
            'diff_cells': 0,
            'added_sheets': [],
            'removed_sheets': [],
            'images_diff': 0
        }

    def run(self):
        """执行全部对比"""
        pythoncom.CoInitialize()
        self.log("正在启动 Excel 应用...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True        # 可见，方便用户查看
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        old_wb = new_wb = None
        try:
            self.progress(5, "打开文件...")
            old_wb = excel.Workbooks.Open(self.old_path, ReadOnly=True)
            new_wb = excel.Workbooks.Open(self.new_path, ReadOnly=True)

            # 1. 比较 Sheet 结构
            self._compare_sheets(old_wb, new_wb)

            # 2. 比较共有 Sheet
            old_sheets = {s.Name: s for s in old_wb.Worksheets}
            new_sheets = {s.Name: s for s in new_wb.Worksheets}
            common_names = set(old_sheets.keys()) & set(new_sheets.keys())
            total = len(common_names)
            for idx, name in enumerate(common_names, 1):
                self.progress(10 + int(80 * idx / total), f"对比 Sheet: {name} ({idx}/{total})")
                self.log(f"正在处理 {name}")
                old_ws = old_sheets[name]
                new_ws = new_sheets[name]
                self._compare_worksheet(old_ws, new_ws, name)

            self.progress(95, "对比完成")
            self.log(f"发现 {self.stats['diff_cells']} 处单元格差异，{len(self.sheet_diffs)} 处 Sheet 结构差异")
        finally:
            if old_wb:
                old_wb.Close(SaveChanges=False)
            if new_wb:
                new_wb.Close(SaveChanges=False)
            excel.ScreenUpdating = True
            excel.DisplayAlerts = True
            pythoncom.CoUninitialize()

    def _compare_sheets(self, old_wb, new_wb):
        old_names = {s.Name for s in old_wb.Worksheets}
        new_names = {s.Name for s in new_wb.Worksheets}
        added = sorted(new_names - old_names)
        removed = sorted(old_names - new_names)
        self.stats['added_sheets'] = added
        self.stats['removed_sheets'] = removed

        for name in added:
            self.sheet_diffs.append({
                'type': '新增Sheet',
                'name': name,
                'old_name': None,
                'new_name': name,
                'desc': f'Sheet "{name}" 只存在于新版'
            })
        for name in removed:
            self.sheet_diffs.append({
                'type': '删除Sheet',
                'name': name,
                'old_name': name,
                'new_name': None,
                'desc': f'Sheet "{name}" 只存在于旧版'
            })

    def _compare_worksheet(self, old_ws, new_ws, sheet_name):
        # 获取使用区域
        old_used = old_ws.UsedRange
        new_used = new_ws.UsedRange
        old_max_row = old_used.Row + old_used.Rows.Count - 1
        old_max_col = old_used.Column + old_used.Columns.Count - 1
        new_max_row = new_used.Row + new_used.Rows.Count - 1
        new_max_col = new_used.Column + new_used.Columns.Count - 1
        max_row = max(old_max_row, new_max_row)
        max_col = max(old_max_col, new_max_col)

        # 逐单元格对比（为提升性能，按行读取批量属性）
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                old_cell = old_ws.Cells(row, col) if row <= old_max_row and col <= old_max_col else None
                new_cell = new_ws.Cells(row, col) if row <= new_max_row and col <= new_max_col else None
                # 获取单元格地址字符串（如 $A$1）
                addr = win32com.client.GetObject(None, old_ws.Cells(row, col).Address(False, False)) if old_cell else new_ws.Cells(row, col).Address(False, False)
                self.stats['total_cells'] += 1
                diff = self._get_cell_diff(old_cell, new_cell)
                if diff:
                    self.stats['diff_cells'] += 1
                    self.diffs.append({
                        'sheet': sheet_name,
                        'address': addr,
                        'row': row, 'col': col,
                        'type': diff['type'],
                        'desc': diff['desc'],
                        'old_sheet': sheet_name,
                        'new_sheet': sheet_name
                    })

        # 合并单元格差异
        self._compare_merged_cells(old_ws, new_ws, sheet_name)
        # 行高列宽差异
        self._compare_row_col_dimensions(old_ws, new_ws, sheet_name)
        # 图片差异
        self._compare_images(old_ws, new_ws, sheet_name)
        # 条件格式差异（仅记录数量变化）
        self._compare_conditional_formats(old_ws, new_ws, sheet_name)

    def _get_cell_diff(self, old_cell, new_cell):
        """比较两个单元格，返回差异描述或 None"""
        if old_cell is None:
            return {'type': '新增单元格', 'desc': f'旧版无此单元格'}
        if new_cell is None:
            return {'type': '删除单元格', 'desc': f'新版无此单元格'}

        old_formula = old_cell.Formula
        new_formula = new_cell.Formula
        if old_formula != new_formula:
            old_val = old_formula if old_formula else "(空)"
            new_val = new_formula if new_formula else "(空)"
            return {'type': '公式变化', 'desc': f'公式: {old_val} → {new_val}'}

        # 比较格式
        font_diff = self._compare_font(old_cell.Font, new_cell.Font)
        fill_diff = self._compare_interior(old_cell.Interior, new_cell.Interior)
        border_diff = self._compare_borders(old_cell.Borders, new_cell.Borders)
        align_diff = self._compare_alignment(old_cell, new_cell)
        numfmt_diff = old_cell.NumberFormat != new_cell.NumberFormat
        merge_diff = old_cell.MergeCells != new_cell.MergeCells

        diffs = []
        if font_diff: diffs.append(f"字体: {font_diff}")
        if fill_diff: diffs.append(f"填充: {fill_diff}")
        if border_diff: diffs.append(f"边框: {border_diff}")
        if align_diff: diffs.append(f"对齐: {align_diff}")
        if numfmt_diff: diffs.append(f"数字格式: {old_cell.NumberFormat} → {new_cell.NumberFormat}")
        if merge_diff: diffs.append("合并状态不同")

        if diffs:
            return {'type': '格式变化', 'desc': '; '.join(diffs)}
        return None

    def _compare_font(self, old_font, new_font):
        changes = []
        if old_font.Name != new_font.Name: changes.append(f"字体名: {old_font.Name}→{new_font.Name}")
        if old_font.Size != new_font.Size: changes.append(f"大小: {old_font.Size}→{new_font.Size}")
        if old_font.Bold != new_font.Bold: changes.append(f"粗体: {old_font.Bold}→{new_font.Bold}")
        if old_font.Italic != new_font.Italic: changes.append(f"斜体: {old_font.Italic}→{new_font.Italic}")
        if old_font.Underline != new_font.Underline: changes.append(f"下划线: {old_font.Underline}→{new_font.Underline}")
        if old_font.Color != new_font.Color: changes.append(f"颜色: {rgb_to_hex(old_font.Color)}→{rgb_to_hex(new_font.Color)}")
        return '; '.join(changes) if changes else None

    def _compare_interior(self, old_int, new_int):
        if old_int.Color != new_int.Color or old_int.Pattern != new_int.Pattern:
            return f"背景: {rgb_to_hex(old_int.Color)}/{old_int.Pattern} → {rgb_to_hex(new_int.Color)}/{new_int.Pattern}"
        return None

    def _compare_borders(self, old_borders, new_borders):
        parts = []
        for idx, name in [(7, "左"), (8, "上"), (9, "下"), (10, "右")]:
            ob = old_borders.Item(idx)
            nb = new_borders.Item(idx)
            if ob.LineStyle != nb.LineStyle or ob.Color != nb.Color:
                parts.append(f"{name}: {ob.LineStyle}/{rgb_to_hex(ob.Color)} → {nb.LineStyle}/{rgb_to_hex(nb.Color)}")
        return '; '.join(parts) if parts else None

    def _compare_alignment(self, old_cell, new_cell):
        if (old_cell.HorizontalAlignment != new_cell.HorizontalAlignment or
            old_cell.VerticalAlignment != new_cell.VerticalAlignment or
            old_cell.WrapText != new_cell.WrapText):
            return (f"水平: {old_cell.HorizontalAlignment}→{new_cell.HorizontalAlignment}, "
                    f"垂直: {old_cell.VerticalAlignment}→{new_cell.VerticalAlignment}, "
                    f"换行: {old_cell.WrapText}→{new_cell.WrapText}")
        return None

    def _compare_merged_cells(self, old_ws, new_ws, sheet_name):
        # 简单比较合并区域数量（更详细需要遍历每个区域，这里简化）
        old_areas = []
        new_areas = []
        try:
            for area in old_ws.UsedRange.MergeAreas:
                old_areas.append(area.Address)
        except:
            pass
        try:
            for area in new_ws.UsedRange.MergeAreas:
                new_areas.append(area.Address)
        except:
            pass
        added = set(new_areas) - set(old_areas)
        removed = set(old_areas) - set(new_areas)
        for addr in added:
            self.diffs.append({
                'sheet': sheet_name,
                'address': addr.split(':')[0],
                'type': '合并单元格新增',
                'desc': f'新增合并区域: {addr}',
                'old_sheet': sheet_name,
                'new_sheet': sheet_name
            })
        for addr in removed:
            self.diffs.append({
                'sheet': sheet_name,
                'address': addr.split(':')[0],
                'type': '合并单元格删除',
                'desc': f'删除合并区域: {addr}',
                'old_sheet': sheet_name,
                'new_sheet': sheet_name
            })

    def _compare_row_col_dimensions(self, old_ws, new_ws, sheet_name):
        # 行高 (简化处理，只比较已使用范围)
        try:
            old_rows = old_ws.Rows
            new_rows = new_ws.Rows
            for r in range(1, old_ws.UsedRange.Rows.Count + 1):
                oh = old_rows(r).RowHeight
                nh = new_rows(r).RowHeight if r <= new_ws.UsedRange.Rows.Count else None
                if oh != nh:
                    self.diffs.append({
                        'sheet': sheet_name,
                        'address': f"A{r}",
                        'type': '行高变化',
                        'desc': f'行高: {oh} → {nh}',
                        'old_sheet': sheet_name,
                        'new_sheet': sheet_name
                    })
        except:
            pass
        # 列宽
        try:
            old_cols = old_ws.Columns
            new_cols = new_ws.Columns
            for c in range(1, old_ws.UsedRange.Columns.Count + 1):
                ow = old_cols(c).ColumnWidth
                nw = new_cols(c).ColumnWidth if c <= new_ws.UsedRange.Columns.Count else None
                if ow != nw:
                    col_letter = win32com.client.GetObject(None, old_ws.Cells(1, c).Address(False, False)).replace("$", "").replace("1", "")
                    self.diffs.append({
                        'sheet': sheet_name,
                        'address': f"{col_letter}1",
                        'type': '列宽变化',
                        'desc': f'列宽: {ow} → {nw}',
                        'old_sheet': sheet_name,
                        'new_sheet': sheet_name
                    })
        except:
            pass

    def _compare_images(self, old_ws, new_ws, sheet_name):
        try:
            old_shapes = old_ws.Shapes
            new_shapes = new_ws.Shapes
            old_imgs = []
            new_imgs = []
            for s in old_shapes:
                if s.Type == 13:  # msoPicture
                    old_imgs.append((s.TopLeftCell.Address(False, False), s.Width, s.Height))
            for s in new_shapes:
                if s.Type == 13:
                    new_imgs.append((s.TopLeftCell.Address(False, False), s.Width, s.Height))
            if old_imgs != new_imgs:
                self.stats['images_diff'] += 1
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': old_imgs[0][0] if old_imgs else (new_imgs[0][0] if new_imgs else 'A1'),
                    'type': '图片差异',
                    'desc': f'旧版图片: {len(old_imgs)}张, 新版: {len(new_imgs)}张, 尺寸变化',
                    'old_sheet': sheet_name,
                    'new_sheet': sheet_name
                })
        except:
            pass

    def _compare_conditional_formats(self, old_ws, new_ws, sheet_name):
        try:
            old_count = old_ws.UsedRange.FormatConditions.Count
            new_count = new_ws.UsedRange.FormatConditions.Count
            if old_count != new_count:
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': 'A1',
                    'type': '条件格式数量变化',
                    'desc': f'条件格式数量: {old_count} → {new_count}',
                    'old_sheet': sheet_name,
                    'new_sheet': sheet_name
                })
        except:
            pass

# ---------------------------- GUI ----------------------------
class DiffViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异对比工具 (原生 Excel 导航)")
        self.root.geometry("950x650")
        self.excel_app = None
        self.old_wb = None
        self.new_wb = None

        # 顶部文件选择
        top_frame = ttk.Frame(root, padding=10)
        top_frame.pack(fill='x')
        ttk.Label(top_frame, text="旧版文件:").grid(row=0, column=0, sticky='w')
        self.old_path = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.old_path, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(top_frame, text="浏览", command=lambda: self.browse(self.old_path)).grid(row=0, column=2)
        ttk.Label(top_frame, text="新版文件:").grid(row=1, column=0, sticky='w', pady=5)
        self.new_path = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.new_path, width=60).grid(row=1, column=1, padx=5)
        ttk.Button(top_frame, text="浏览", command=lambda: self.browse(self.new_path)).grid(row=1, column=2)

        btn_frame = ttk.Frame(top_frame)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=5)
        self.start_btn = ttk.Button(btn_frame, text="开始对比", command=self.start_compare)
        self.start_btn.pack(side='left', padx=5)
        self.jump_btn = ttk.Button(btn_frame, text="跳转到选中项", command=self.jump_to_selected, state='disabled')
        self.jump_btn.pack(side='left', padx=5)

        # 进度条
        self.progress = ttk.Progressbar(root, mode='determinate')
        self.progress.pack(fill='x', padx=10)

        # 主面板：左侧树形差异，右侧详情
        paned = ttk.PanedWindow(root, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=10, pady=5)

        # 左侧树
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=3)
        self.tree = ttk.Treeview(left_frame, columns=('address', 'type'), show='tree headings')
        self.tree.heading('#0', text='Sheet / 差异项')
        self.tree.heading('address', text='位置')
        self.tree.heading('type', text='类型')
        self.tree.column('#0', width=200)
        self.tree.column('address', width=80)
        self.tree.column('type', width=120)
        scroll_tree = ttk.Scrollbar(left_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_tree.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scroll_tree.pack(side='right', fill='y')
        self.tree.bind('<Double-1>', self.on_tree_double_click)

        # 右侧详情
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=2)
        ttk.Label(right_frame, text="差异详情", font=('微软雅黑', 10, 'bold')).pack(anchor='w')
        self.detail_text = tk.Text(right_frame, wrap='word', height=20)
        self.detail_text.pack(fill='both', expand=True)

        # 日志
        log_frame = ttk.LabelFrame(root, text="日志", padding=5)
        log_frame.pack(fill='x', padx=10, pady=(0, 10))
        self.log_text = tk.Text(log_frame, height=6, wrap='word')
        self.log_text.pack(fill='both', expand=True)

        self.comparer = None
        self.diff_items = []   # 存储所有差异条目对应的树节点和完整数据

    def browse(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if path:
            var.set(path)

    def log(self, msg):
        self.log_text.insert('end', f"{time.strftime('%H:%M:%S')} {msg}\n")
        self.log_text.see('end')
        self.root.update_idletasks()

    def update_progress(self, value, status=""):
        self.progress['value'] = value
        if status:
            self.log(status)
        self.root.update_idletasks()

    def start_compare(self):
        old = self.old_path.get()
        new = self.new_path.get()
        if not old or not new:
            messagebox.showerror("错误", "请选择两个 Excel 文件")
            return
        self.start_btn.configure(state='disabled')
        self.tree.delete(*self.tree.get_children())
        self.detail_text.delete('1.0', 'end')
        self.diff_items = []

        def run():
            try:
                self.comparer = ExcelComparer(old, new, self.log, self.update_progress)
                self.comparer.run()
                self.root.after(0, self.populate_tree)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
            finally:
                self.root.after(0, lambda: self.start_btn.configure(state='normal'))

        threading.Thread(target=run, daemon=True).start()

    def populate_tree(self):
        """将差异结果填充到树形控件"""
        diffs_by_sheet = {}
        for diff in self.comparer.diffs:
            sheet = diff['sheet']
            if sheet not in diffs_by_sheet:
                diffs_by_sheet[sheet] = []
            diffs_by_sheet[sheet].append(diff)

        # 添加 Sheet 结构差异
        if self.comparer.sheet_diffs:
            struct_node = self.tree.insert('', 'end', text='📋 Sheet 结构差异', open=True)
            for sd in self.comparer.sheet_diffs:
                node = self.tree.insert(struct_node, 'end',
                                        text=sd['desc'],
                                        values=(sd['name'], sd['type']))
                self.diff_items.append((node, {'type': 'sheet_struct', 'data': sd}))

        # 按 Sheet 分组
        for sheet_name, items in sorted(diffs_by_sheet.items()):
            sheet_node = self.tree.insert('', 'end', text=f"📄 {sheet_name}", open=True)
            for diff in items:
                node = self.tree.insert(sheet_node, 'end',
                                        text=diff['desc'][:80],   # 简要描述
                                        values=(diff['address'], diff['type']))
                self.diff_items.append((node, {'type': 'cell', 'data': diff}))

        self.jump_btn.configure(state='normal')
        self.log("树形列表已加载，双击可跳转")

    def on_tree_double_click(self, event):
        self.jump_to_selected()

    def jump_to_selected(self):
        selection = self.tree.selection()
        if not selection:
            return
        node = selection[0]
        # 查找对应的数据
        target = None
        for n, data in self.diff_items:
            if n == node:
                target = data
                break
        if not target:
            return

        # 显示详情
        if target['type'] == 'cell':
            diff = target['data']
            detail = (f"Sheet: {diff['sheet']}\n"
                      f"单元格: {diff['address']}\n"
                      f"类型: {diff['type']}\n"
                      f"描述: {diff['desc']}")
            self.detail_text.delete('1.0', 'end')
            self.detail_text.insert('1.0', detail)

            # COM 跳转
            self._navigate_to_cell(diff['sheet'], diff['address'])
        elif target['type'] == 'sheet_struct':
            sd = target['data']
            detail = f"类型: {sd['type']}\n描述: {sd['desc']}"
            self.detail_text.delete('1.0', 'end')
            self.detail_text.insert('1.0', detail)
            # 跳转到 A1
            if sd['new_name']:
                self._navigate_to_cell(sd['new_name'], 'A1')
            elif sd['old_name']:
                self._navigate_to_cell(sd['old_name'], 'A1')

    def _navigate_to_cell(self, sheet_name, address):
        """通过 COM 在新旧两个 Excel 中选中对应单元格"""
        try:
            # 确保 Excel 应用已打开，并保持两个工作簿
            if not self.excel_app:
                pythoncom.CoInitialize()
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                self.excel_app.Visible = True
                self.old_wb = self.excel_app.Workbooks.Open(self.old_path.get(), ReadOnly=True)
                self.new_wb = self.excel_app.Workbooks.Open(self.new_path.get(), ReadOnly=True)

            # 激活新旧两个窗口，选择对应 Sheet 和单元格
            for wb, desc in [(self.old_wb, "旧版"), (self.new_wb, "新版")]:
                try:
                    ws = wb.Worksheets(sheet_name)
                    ws.Activate()
                    ws.Range(address).Select()
                    self.log(f"{desc} 已跳转到 {sheet_name}!{address}")
                except Exception as e:
                    self.log(f"{desc} 跳转失败: {e}")
        except Exception as e:
            messagebox.showerror("跳转错误", f"无法操作 Excel: {e}")

# ---------------------------- 主入口 ----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = DiffViewer(root)
    root.mainloop()
