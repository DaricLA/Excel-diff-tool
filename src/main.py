"""
Excel 差异对比工具（优化版）
- 可拖动分隔条，左侧列表初始占2/3，右侧详情占1/3
- 增强富文本差异检测
- 使用Goto定位单元格，视觉更醒目
- 完全由用户手动打开Excel，工具仅检测与跳转
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
from openpyxl.utils import get_column_letter

# ---------------------------- 辅助函数 ----------------------------
def rgb_to_hex(rgb):
    if rgb is None or rgb.rgb is None:
        return "None"
    if isinstance(rgb.rgb, str):
        return rgb.rgb
    return str(rgb.rgb)[2:] if len(str(rgb.rgb)) == 8 else str(rgb.rgb)

def cell_address(col_idx, row_idx):
    return f"{get_column_letter(col_idx)}{row_idx}"

def normalize_path(path):
    try:
        return os.path.normpath(os.path.realpath(path))
    except:
        return os.path.normpath(path)

# ---------------------------- 对比引擎 ----------------------------
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
        old_wb = load_workbook(self.old_path, data_only=False)
        self.progress(15, "加载新版文件...")
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
        old_val = old_cell.value
        new_val = new_cell.value
        # 1. 值/公式/富文本差异
        if not self._value_equal(old_val, new_val):
            old_str = f"{old_val}" if old_val is not None else "(空)"
            new_str = f"{new_val}" if new_val is not None else "(空)"
            # 如果是富文本，生成详细的格式变更描述
            if isinstance(old_val, CellRichText) or isinstance(new_val, CellRichText):
                rich_desc = self._rich_text_diff(old_val, new_val)
                return {'type': '内容(含富文本)', 'desc': rich_desc}
            elif (isinstance(old_val, str) and old_val.startswith('=')) or \
                 (isinstance(new_val, str) and new_val.startswith('=')):
                return {'type': '公式变化', 'desc': f'{old_str} → {new_str}'}
            else:
                return {'type': '值变化', 'desc': f'{old_str} → {new_str}'}

        # 2. 格式差异（整体字体、填充等）
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
            self.diffs.append({'sheet':sheet_name,'address':addr.split(':')[0],'type':'合并新增','desc':f'新增合并区域 {addr}'})
        for addr in old_merged - new_merged:
            self.diffs.append({'sheet':sheet_name,'address':addr.split(':')[0],'type':'合并删除','desc':f'删除合并区域 {addr}'})

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

    def _get_images_from_ws(self, ws):
        images = []
        if hasattr(ws, '_images') and ws._images:
            for img in ws._images:
                try:
                    anchor = img.anchor
                    if anchor and hasattr(anchor, '_from'):
                        col = anchor._from.col
                        row = anchor._from.row
                        addr = cell_address(col, row)
                        images.append((addr, img.width, img.height))
                except:
                    pass
        if images:
            return sorted(images, key=lambda x: x[0])
        if hasattr(ws, '_drawing') and ws._drawing:
            for anchor in ws._drawing.anchors:
                if hasattr(anchor, 'image'):
                    img = anchor.image
                    col = anchor._from.col if hasattr(anchor, '_from') else 0
                    row = anchor._from.row if hasattr(anchor, '_from') else 0
                    addr = cell_address(col, row)
                    images.append((addr, img.width, img.height))
        return sorted(images, key=lambda x: x[0])

    def _compare_images(self, old_ws, new_ws, sheet_name):
        old_imgs = self._get_images_from_ws(old_ws)
        new_imgs = self._get_images_from_ws(new_ws)
        if not old_imgs and not new_imgs:
            return
        diff_detected = False
        if len(old_imgs) != len(new_imgs):
            diff_detected = True
        else:
            for (addr1, w1, h1), (addr2, w2, h2) in zip(old_imgs, new_imgs):
                if addr1 != addr2 or abs(w1 - w2) > 0.1 or abs(h1 - h2) > 0.1:
                    diff_detected = True
                    break
        if diff_detected:
            self.stats['images_diff'] += 1
            anchor = old_imgs[0][0] if old_imgs else (new_imgs[0][0] if new_imgs else 'A1')
            desc_parts = [f"旧版 {len(old_imgs)} 张，新版 {len(new_imgs)} 张"]
            if len(old_imgs) != len(new_imgs):
                desc_parts.append("图片数量不同")
            else:
                desc_parts.append("图片尺寸或位置变化")
            self.diffs.append({
                'sheet': sheet_name,
                'address': anchor,
                'type': '图片差异',
                'desc': '；'.join(desc_parts)
            })

    def _compare_conditional_formats(self, old_ws, new_ws, sheet_name):
        old_cfs = list(old_ws.conditional_formatting)
        new_cfs = list(new_ws.conditional_formatting)
        if not old_cfs and not new_cfs:
            return
        def cf_to_dict(cf):
            rules = []
            for rule in cf.rules:
                rules.append({
                    'type': rule.type,
                    'priority': rule.priority,
                    'formula': str(rule.formula),
                    'dxf': str(rule.dxf) if rule.dxf else None
                })
            return {'sqref': str(cf.sqref), 'rules': rules}
        old_cf_dicts = [cf_to_dict(cf) for cf in old_cfs]
        new_cf_dicts = [cf_to_dict(cf) for cf in new_cfs]
        if old_cf_dicts != new_cf_dicts:
            self.diffs.append({
                'sheet': sheet_name,
                'address': 'A1',
                'type': '条件格式变化',
                'desc': f'条件格式规则有差异 (旧版 {len(old_cfs)} 条, 新版 {len(new_cfs)} 条)'
            })

# ---------------------------- GUI + COM 跳转 ----------------------------
class DiffViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异对比工具 (手动打开文件)")
        self.root.geometry("950x650")
        self.root.option_add("*Font", ("微软雅黑", 10))
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()

        # 顶部文件选择区域
        top = ttk.Frame(root, padding=10)
        top.grid(row=0, column=0, sticky='ew')
        ttk.Label(top, text="旧版文件:", font=("微软雅黑", 10)).grid(row=0,column=0,sticky='w')
        ttk.Entry(top, textvariable=self.old_path, width=60).grid(row=0,column=1,padx=5)
        ttk.Button(top, text="浏览", command=lambda: self.browse(self.old_path)).grid(row=0,column=2)
        ttk.Label(top, text="新版文件:", font=("微软雅黑", 10)).grid(row=1,column=0,sticky='w',pady=5)
        ttk.Entry(top, textvariable=self.new_path, width=60).grid(row=1,column=1,padx=5)
        ttk.Button(top, text="浏览", command=lambda: self.browse(self.new_path)).grid(row=1,column=2)

        btn_frame = ttk.Frame(top)
        btn_frame.grid(row=2,column=0,columnspan=3,pady=5)
        self.start_btn = tk.Button(btn_frame, text="开始对比", font=("微软雅黑", 11, "bold"),
                                   bg="#0078D7", fg="white", width=12, command=self.start_compare)
        self.start_btn.pack(side='left',padx=5)
        self.jump_btn = tk.Button(btn_frame, text="跳转到选中项", font=("微软雅黑", 10),
                                  command=self.jump_to_selected, state='disabled')
        self.jump_btn.pack(side='left',padx=5)

        self.progress = ttk.Progressbar(root, mode='determinate')
        self.progress.grid(row=1, column=0, sticky='ew', padx=10, pady=(5,0))

        # 可拖动的 PanedWindow：左侧列表，右侧详情
        paned = ttk.PanedWindow(root, orient='horizontal')
        paned.grid(row=2, column=0, sticky='nsew', padx=10, pady=5)
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(0, weight=1)

        left_frame = ttk.Frame(paned)
        right_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=2)   # 初始比例 2/3
        paned.add(right_frame, weight=1)  # 初始比例 1/3

        # 左侧列表
        self.tree = ttk.Treeview(left_frame, columns=('address','type'), show='tree headings')
        self.tree.heading('#0',text='Sheet / 差异项')
        self.tree.heading('address',text='位置')
        self.tree.heading('type',text='类型')
        self.tree.column('#0',width=200); self.tree.column('address',width=80); self.tree.column('type',width=120)
        scroll_y = ttk.Scrollbar(left_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Double-1>', self.jump_to_selected)

        # 右侧详情
        ttk.Label(right_frame, text="差异详情", font=('微软雅黑',10,'bold')).pack(anchor='w')
        self.detail = tk.Text(right_frame, wrap='word', height=20, font=("微软雅黑", 10))
        self.detail.pack(fill='both',expand=True)

        # 日志区域
        logf = ttk.LabelFrame(root, text="日志", padding=5)
        logf.grid(row=3, column=0, sticky='ew', padx=10, pady=(0,10))
        self.log_text = tk.Text(logf, height=8, wrap='word', font=("微软雅黑", 10))
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
                self.root.after(0, self.on_comparison_finished)

        threading.Thread(target=worker, daemon=True).start()

    def on_comparison_finished(self):
        self.start_btn.configure(state='normal')
        self.progress['value'] = 100

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
        self.log("树形列表已加载，单击查看详情，双击跳转")

    def on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        node = sel[0]
        target = None
        for n, d in self.diff_items:
            if n == node:
                target = d; break
        if not target:
            return
        if target['type'] == 'cell':
            d = target['data']
            self.detail.delete('1.0','end')
            self.detail.insert('1.0', f"Sheet: {d['sheet']}\n单元格: {d['address']}\n类型: {d['type']}\n描述: {d['desc']}")
        elif target['type'] == 'sheet_struct':
            sd = target['data']
            self.detail.delete('1.0','end')
            self.detail.insert('1.0', f"类型: {sd['type']}\n描述: {sd['desc']}")

    def jump_to_selected(self, event=None):
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
            sheet = d['sheet']
            address = d['address']
        elif target['type'] == 'sheet_struct':
            sd = target['data']
            sheet = sd.get('name', sd.get('new_name'))
            address = 'A1'
        else:
            return

        def navigate():
            pythoncom.CoInitialize()
            try:
                try:
                    excel = win32com.client.GetObject(Class="Excel.Application")
                except:
                    messagebox.showinfo("提示", "未检测到正在运行的 Excel 进程，请手动打开两个文件后再跳转。")
                    return

                old_path = normalize_path(self.old_path.get())
                new_path = normalize_path(self.new_path.get())
                old_wb = None
                new_wb = None
                for wb in excel.Workbooks:
                    path = normalize_path(wb.FullName)
                    if path == old_path:
                        old_wb = wb
                    elif path == new_path:
                        new_wb = wb

                if not old_wb or not new_wb:
                    missing = []
                    if not old_wb: missing.append(f"旧版文件: {self.old_path.get()}")
                    if not new_wb: missing.append(f"新版文件: {self.new_path.get()}")
                    messagebox.showinfo("文件未打开", "以下文件未在 Excel 中打开：\n" + "\n".join(missing) + "\n请手动打开后重试。")
                    return

                # 使用 Goto 定位并居中，视觉效果更醒目
                for wb, desc in [(old_wb, "旧版"), (new_wb, "新版")]:
                    try:
                        ws = wb.Worksheets(sheet)
                        ws.Activate()
                        excel.Goto(ws.Range(address), Scroll=True)  # 关键：使用 Goto
                        self.log(f"{desc} 已跳转到 {sheet}!{address}")
                    except Exception as e:
                        self.log(f"{desc} 跳转失败: {e}")

                excel.Visible = True
                excel.WindowState = -4137  # xlMaximized
                self.root.after(100, lambda: self.root.lower())
            except Exception as e:
                self.log(f"跳转出错: {e}")
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=navigate, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = DiffViewer(root)
    root.mainloop()
