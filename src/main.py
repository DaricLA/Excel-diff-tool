"""
Excel 差异对比工具（openpyxl 高速对比 + COM 跳转）
- 使用 openpyxl 快速读取新旧 Excel 的所有信息
- 支持单元格值、公式、字体、填充、边框、对齐、数字格式、合并单元格
- 支持富文本（CellRichText）逐段比较
- 支持图片数量及尺寸变化检测
- 支持条件格式数量变化检测
- 双击差异项，自动打开 Excel 并跳转到对应单元格
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import os
import pythoncom
import win32com.client
from win32com.client import constants
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText

# ---------------------------- 辅助函数 ----------------------------
def rgb_to_hex(rgb):
    """将 openpyxl 的颜色对象转换为 #RRGGBB 字符串"""
    if rgb is None or rgb.rgb is None:
        return "None"
    if isinstance(rgb.rgb, str):
        return rgb.rgb
    # 如果 rgb 是 ARGB 字符串，去掉前两位透明度
    return str(rgb.rgb)[2:] if len(str(rgb.rgb)) == 8 else str(rgb.rgb)

def cell_address(col_idx, row_idx):
    """将列、行索引转换为 'A1' 格式"""
    from openpyxl.utils import get_column_letter
    return f"{get_column_letter(col_idx)}{row_idx}"

# ---------------------------- openpyxl 对比引擎 ----------------------------
class OpenpyxlComparer:
    def __init__(self, old_path, new_path, log_callback=None, progress_callback=None):
        self.old_path = old_path
        self.new_path = new_path
        self.log = log_callback if log_callback else print
        self.progress = progress_callback if progress_callback else lambda v,s: None
        self.diffs = []
        self.sheet_diffs = []
        self.stats = {
            'total_cells': 0, 'diff_cells': 0,
            'added_sheets': [], 'removed_sheets': [], 'images_diff': 0
        }

    def run(self):
        self.progress(5, "加载旧版文件...")
        self.log("正在使用 openpyxl 加载旧版文件...")
        old_wb = load_workbook(self.old_path, data_only=False)
        self.progress(15, "加载新版文件...")
        self.log("正在加载新版文件...")
        new_wb = load_workbook(self.new_path, data_only=False)

        self._compare_sheets(old_wb, new_wb)
        common = set(old_wb.sheetnames) & set(new_wb.sheetnames)
        total = len(common)
        for idx, sheet_name in enumerate(common, 1):
            self.progress(20 + int(70 * idx / total), f"对比 {sheet_name}...")
            self.log(f"正在对比 {sheet_name} ({idx}/{total})")
            old_ws = old_wb[sheet_name]
            new_ws = new_wb[sheet_name]
            self._compare_worksheet(old_ws, new_ws, sheet_name)

        self.progress(95, "对比完成")
        self.log(f"发现 {self.stats['diff_cells']} 处单元格差异，{len(self.sheet_diffs)} 处 Sheet 差异")
        return True

    def _compare_sheets(self, old_wb, new_wb):
        old_names = set(old_wb.sheetnames)
        new_names = set(new_wb.sheetnames)
        self.stats['added_sheets'] = sorted(new_names - old_names)
        self.stats['removed_sheets'] = sorted(old_names - new_names)
        for name in self.stats['added_sheets']:
            self.sheet_diffs.append({'type':'新增Sheet', 'name':name, 'desc':f'Sheet "{name}" 只存在于新版'})
        for name in self.stats['removed_sheets']:
            self.sheet_diffs.append({'type':'删除Sheet', 'name':name, 'desc':f'Sheet "{name}" 只存在于旧版'})

    def _compare_worksheet(self, old_ws, new_ws, sheet_name):
        max_row = max(old_ws.max_row, new_ws.max_row)
        max_col = max(old_ws.max_column, new_ws.max_column)
        self.stats['total_cells'] += max_row * max_col

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                old_cell = old_ws.cell(row, col)
                new_cell = new_ws.cell(row, col)
                diff = self._get_cell_diff(old_cell, new_cell)
                if diff:
                    self.stats['diff_cells'] += 1
                    self.diffs.append({
                        'sheet': sheet_name,
                        'address': cell_address(col, row),
                        'type': diff['type'],
                        'desc': diff['desc']
                    })

        self._compare_merged_cells(old_ws, new_ws, sheet_name)
        self._compare_row_col_dimensions(old_ws, new_ws, sheet_name)
        self._compare_images(old_ws, new_ws, sheet_name)
        self._compare_conditional_formats(old_ws, new_ws, sheet_name)

    def _get_cell_diff(self, old_cell, new_cell):
        # 值 / 公式变化
        old_val = old_cell.value
        new_val = new_cell.value
        if not self._value_equal(old_val, new_val):
            old_str = f"{old_val}" if old_val is not None else "(空)"
            new_str = f"{new_val}" if new_val is not None else "(空)"
            # 富文本详细比较
            if isinstance(old_val, CellRichText) or isinstance(new_val, CellRichText):
                rich_desc = self._rich_text_diff(old_val, new_val)
                return {'type': '内容(含富文本)', 'desc': rich_desc}
            elif (isinstance(old_val, str) and old_val.startswith('=')) or \
                 (isinstance(new_val, str) and new_val.startswith('=')):
                return {'type': '公式变化', 'desc': f'{old_str} → {new_str}'}
            else:
                return {'type': '值变化', 'desc': f'{old_str} → {new_str}'}

        # 格式比较
        descs = []
        font_diff = self._cmp_font(old_cell.font, new_cell.font)
        if font_diff: descs.append(f"字体: {font_diff}")
        fill_diff = self._cmp_fill(old_cell.fill, new_cell.fill)
        if fill_diff: descs.append(f"填充: {fill_diff}")
        border_diff = self._cmp_border(old_cell.border, new_cell.border)
        if border_diff: descs.append(f"边框: {border_diff}")
        align_diff = self._cmp_alignment(old_cell.alignment, new_cell.alignment)
        if align_diff: descs.append(f"对齐: {align_diff}")
        num_diff = old_cell.number_format != new_cell.number_format
        if num_diff: descs.append(f"数字格式: {old_cell.number_format} → {new_cell.number_format}")
        if descs:
            return {'type': '格式变化', 'desc': '; '.join(descs)}
        return None

    def _value_equal(self, v1, v2):
        if type(v1) != type(v2):
            return False
        if isinstance(v1, CellRichText) and isinstance(v2, CellRichText):
            if len(v1) != len(v2):
                return False
            for t1, t2 in zip(v1, v2):
                if t1.text != t2.text or not self._font_equal(t1.font, t2.font):
                    return False
            return True
        return v1 == v2

    def _rich_text_diff(self, old_rt, new_rt):
        """生成富文本差异描述"""
        if not isinstance(old_rt, CellRichText):
            old_rt = CellRichText(str(old_rt) if old_rt else "")
        if not isinstance(new_rt, CellRichText):
            new_rt = CellRichText(str(new_rt) if new_rt else "")
        old_plain = str(old_rt)
        new_plain = str(new_rt)
        if old_plain != new_plain:
            return f"内容(含富文本): {old_plain} → {new_plain}"
        lines = []
        for i, (t1, t2) in enumerate(zip(old_rt, new_rt)):
            if t1.text != t2.text or not self._font_equal(t1.font, t2.font):
                seg = t1.text if t1.text else "(空)"
                changes = []
                if t1.font.name != t2.font.name: changes.append(f"字体: {t1.font.name}→{t2.font.name}")
                if t1.font.size != t2.font.size: changes.append(f"大小: {t1.font.size}→{t2.font.size}")
                if t1.font.bold != t2.font.bold: changes.append(f"粗: {t1.font.bold}→{t2.font.bold}")
                if t1.font.italic != t2.font.italic: changes.append(f"斜: {t1.font.italic}→{t2.font.italic}")
                if t1.font.color != t2.font.color: changes.append(f"颜色: {rgb_to_hex(t1.font.color)}→{rgb_to_hex(t2.font.color)}")
                if changes:
                    lines.append(f"段{i+1} '{seg}': {'; '.join(changes)}")
                else:
                    lines.append(f"段{i+1} '{seg}': 变为 '{t2.text}'")
        return "富文本格式变更:\n" + "\n".join(lines) if lines else "富文本格式有细微变化"

    # ---------- 格式比较方法 ----------
    @staticmethod
    def _font_equal(f1, f2):
        return (f1.name == f2.name and f1.size == f2.size and f1.bold == f2.bold and
                f1.italic == f2.italic and f1.underline == f2.underline and f1.color == f2.color)

    def _cmp_font(self, f1, f2):
        changes = []
        if f1.name != f2.name: changes.append(f"字体名: {f1.name}→{f2.name}")
        if f1.size != f2.size: changes.append(f"大小: {f1.size}→{f2.size}")
        if f1.bold != f2.bold: changes.append(f"粗体: {f1.bold}→{f2.bold}")
        if f1.italic != f2.italic: changes.append(f"斜体: {f1.italic}→{f2.italic}")
        if f1.underline != f2.underline: changes.append(f"下划线: {f1.underline}→{f2.underline}")
        if f1.color != f2.color: changes.append(f"颜色: {rgb_to_hex(f1.color)}→{rgb_to_hex(f2.color)}")
        return '; '.join(changes) if changes else None

    def _cmp_fill(self, f1, f2):
        if f1.fill_type != f2.fill_type or f1.start_color != f2.start_color or f1.end_color != f2.end_color:
            return f"类型: {f1.fill_type}→{f2.fill_type}, 颜色: {rgb_to_hex(f1.start_color)}→{rgb_to_hex(f2.start_color)}"
        return None

    def _cmp_border(self, b1, b2):
        parts = []
        for side in ['left', 'right', 'top', 'bottom']:
            s1 = getattr(b1, side)
            s2 = getattr(b2, side)
            if s1.style != s2.style or s1.color != s2.color:
                parts.append(f"{side}: {s1.style}/{rgb_to_hex(s1.color)}→{s2.style}/{rgb_to_hex(s2.color)}")
        return '; '.join(parts) if parts else None

    def _cmp_alignment(self, a1, a2):
        changes = []
        if a1.horizontal != a2.horizontal: changes.append(f"水平: {a1.horizontal}→{a2.horizontal}")
        if a1.vertical != a2.vertical: changes.append(f"垂直: {a1.vertical}→{a2.vertical}")
        if a1.wrap_text != a2.wrap_text: changes.append(f"自动换行: {a1.wrap_text}→{a2.wrap_text}")
        return '; '.join(changes) if changes else None

    def _compare_merged_cells(self, old_ws, new_ws, sheet_name):
        old_merged = set(str(m) for m in old_ws.merged_cells.ranges)
        new_merged = set(str(m) for m in new_ws.merged_cells.ranges)
        for addr in new_merged - old_merged:
            start = addr.split(':')[0]
            self.diffs.append({'sheet':sheet_name,'address':start,'type':'合并新增','desc':f'新增合并区域 {addr}'})
        for addr in old_merged - new_merged:
            start = addr.split(':')[0]
            self.diffs.append({'sheet':sheet_name,'address':start,'type':'合并删除','desc':f'删除合并区域 {addr}'})

    def _compare_row_col_dimensions(self, old_ws, new_ws, sheet_name):
        for row_idx, rd in old_ws.row_dimensions.items():
            oh = rd.height
            nh = new_ws.row_dimensions[row_idx].height if row_idx in new_ws.row_dimensions else None
            if oh != nh:
                self.diffs.append({'sheet':sheet_name,'address':f"A{row_idx}",'type':'行高变化','desc':f'行高: {oh} → {nh}'})
        for row_idx, rd in new_ws.row_dimensions.items():
            if row_idx not in old_ws.row_dimensions:
                nh = rd.height
                if nh is not None:
                    self.diffs.append({'sheet':sheet_name,'address':f"A{row_idx}",'type':'行高新设置','desc':f'行高: {nh}'})

        for col_letter, cd in old_ws.column_dimensions.items():
            ow = cd.width
            nw = new_ws.column_dimensions[col_letter].width if col_letter in new_ws.column_dimensions else None
            if ow != nw:
                from openpyxl.utils import column_index_from_string
                col_idx = column_index_from_string(col_letter)
                addr = cell_address(col_idx, 1)
                self.diffs.append({'sheet':sheet_name,'address':addr,'type':'列宽变化','desc':f'列宽({col_letter}): {ow} → {nw}'})
        for col_letter, cd in new_ws.column_dimensions.items():
            if col_letter not in old_ws.column_dimensions:
                nw = cd.width
                if nw is not None:
                    from openpyxl.utils import column_index_from_string
                    col_idx = column_index_from_string(col_letter)
                    addr = cell_address(col_idx, 1)
                    self.diffs.append({'sheet':sheet_name,'address':addr,'type':'列宽新设置','desc':f'列宽({col_letter}): {nw}'})

    def _compare_images(self, old_ws, new_ws, sheet_name):
        """比较图片：通过 openpyxl 的 drawing 获取图片信息"""
        def get_images(ws):
            imgs = []
            try:
                if ws._drawing:
                    for anchor in ws._drawing.anchors:
                        if hasattr(anchor, 'image'):
                            img = anchor.image
                            col = anchor._from.col
                            row = anchor._from.row
                            imgs.append((cell_address(col, row), img.width, img.height))
            except:
                pass
            return imgs
        old_imgs = get_images(old_ws)
        new_imgs = get_images(new_ws)
        if old_imgs != new_imgs:
            self.stats['images_diff'] += 1
            anchor = old_imgs[0][0] if old_imgs else (new_imgs[0][0] if new_imgs else 'A1')
            self.diffs.append({'sheet':sheet_name,'address':anchor,'type':'图片差异','desc':f'图片数量/尺寸变化'})

    def _compare_conditional_formats(self, old_ws, new_ws, sheet_name):
        """比较条件格式数量"""
        old_count = len(old_ws.conditional_formatting)
        new_count = len(new_ws.conditional_formatting)
        if old_count != new_count:
            self.diffs.append({'sheet':sheet_name,'address':'A1','type':'条件格式数量变化','desc':f'条件格式: {old_count} → {new_count}'})

# ---------------------------- GUI + COM 跳转 ----------------------------
class DiffViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异对比工具（高速 + 跳转）")
        self.root.geometry("950x650")
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

        btn_frame = ttk.Frame(top)
        btn_frame.grid(row=2,column=0,columnspan=3,pady=5)
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
        self.result_data = None

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
                comparer = OpenpyxlComparer(old, new, self.log, self.update_progress)
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
            self._jump_to_excel(d['sheet'], d['address'])
        elif target['type'] == 'sheet_struct':
            sd = target['data']
            self.detail.delete('1.0','end')
            self.detail.insert('1.0', f"类型: {sd['type']}\n描述: {sd['desc']}")
            sheet = sd.get('name', sd.get('new_name'))
            if sheet:
                self._jump_to_excel(sheet, 'A1')

    def _jump_to_excel(self, sheet_name, address):
        """使用 COM 打开 Excel 并跳转到指定单元格"""
        def navigate():
            pythoncom.CoInitialize()
            try:
                excel = win32com.client.GetObject(Class="Excel.Application")
            except:
                # 如果没有 Excel 在运行，则启动一个新的
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = True
            excel.DisplayAlerts = False
            old_wb = None
            new_wb = None
            try:
                old_wb = excel.Workbooks.Open(self.old_path.get(), ReadOnly=True)
                new_wb = excel.Workbooks.Open(self.new_path.get(), ReadOnly=True)
                excel.Windows.Arrange(constants.xlArrangeStyleTiled)
            except:
                pass

            for wb, desc in [(old_wb, "旧版"), (new_wb, "新版")]:
                if wb:
                    try:
                        ws = wb.Worksheets(sheet_name)
                        ws.Activate()
                        ws.Range(address).Select()
                        self.log(f"{desc} 已跳转到 {sheet_name}!{address}")
                    except Exception as e:
                        self.log(f"{desc} 跳转失败: {e}")
            pythoncom.CoUninitialize()

        threading.Thread(target=navigate, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = DiffViewer(root)
    root.mainloop()
