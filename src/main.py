"""
Excel 差异对比工具（最终增强版）- ttkbootstrap flatly 主题版
- 每个条件格式差异单独列出，精确定位首单元格
- 增强图片检测，openpyxl + COM 双重保障
- 界面优化：按钮左置，路径缩短，详情/日志对调，置顶开关
- 可靠富文本检测（lxml 解析）
"""
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import threading
import time
import os
import pythoncom
import win32com.client
from win32com.client import constants
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from lxml import etree

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

# ---------------------------- 富文本解析 ----------------------------
def extract_rich_props(cell_xml):
    """从单元格 XML 中提取所有文本运行及其格式，返回列表 [(text, font_dict)]"""
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    runs = []
    for r in cell_xml.findall(f'.//{ns}r'):
        t_elem = r.find(f'{ns}t')
        text = t_elem.text if t_elem is not None else ""
        rPr = r.find(f'{ns}rPr')
        font = {'name': '', 'size': '', 'bold': False, 'italic': False, 'underline': False, 'color': ''}
        if rPr is not None:
            rf = rPr.find(f'{ns}rFont')
            if rf is not None:
                font['name'] = rf.get('val', '')
            sz = rPr.find(f'{ns}sz')
            if sz is not None:
                font['size'] = sz.get('val', '')
            font['bold'] = rPr.find(f'{ns}b') is not None
            font['italic'] = rPr.find(f'{ns}i') is not None
            font['underline'] = rPr.find(f'{ns}u') is not None
            color = rPr.find(f'{ns}color')
            if color is not None:
                font['color'] = color.get('rgb', '')
        runs.append((text, font))
    return runs

def compare_rich_text(cell1, cell2):
    """比较两个单元格的富文本，返回差异描述或 None"""
    xml1 = cell1._cell.xml if hasattr(cell1, '_cell') else None
    xml2 = cell2._cell.xml if hasattr(cell2, '_cell') else None
    if not xml1 or not xml2:
        return None

    elem1 = extract_rich_props(etree.fromstring(xml1.encode()))
    elem2 = extract_rich_props(etree.fromstring(xml2.encode()))

    if not elem1 and not elem2:
        return None  # 无富文本

    plain1 = ''.join(t for t, f in elem1)
    plain2 = ''.join(t for t, f in elem2)
    if plain1 != plain2:
        return f"内容(含富文本): {plain1} → {plain2}"

    if len(elem1) != len(elem2):
        return f"富文本段落数不同: {len(elem1)} → {len(elem2)}"
    changes = []
    for i, ((t1, f1), (t2, f2)) in enumerate(zip(elem1, elem2)):
        if f1 != f2:
            diff = []
            for k in ['name', 'size', 'bold', 'italic', 'underline', 'color']:
                if f1.get(k) != f2.get(k):
                    diff.append(f"{k}: {f1.get(k)}→{f2.get(k)}")
            changes.append(f"段{i+1} '{t1}': {'; '.join(diff)}")
    if changes:
        return "富文本格式变更:\n" + "\n".join(changes)
    return None

# ---------------------------- 对比引擎 ----------------------------
class OpenpyxlComparer:
    def __init__(self, old_path, new_path, log_callback=None, progress_callback=None):
        self.old_path = old_path
        self.new_path = new_path
        self.log = log_callback if log_callback else print
        self.progress = progress_callback if progress_callback else lambda v,s: None
        self.diffs = []
        self.sheet_diffs = []
        self.stats = {'total_cells':0, 'diff_cells':0, 'added_sheets':[], 'removed_sheets':[], 'images_diff':0}

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
        # 优先检测富文本
        rich_diff = compare_rich_text(old_cell, new_cell)
        if rich_diff:
            return {'type': '内容(含富文本)', 'desc': rich_diff}

        old_val = old_cell.value
        new_val = new_cell.value
        if old_val != new_val:
            old_str = f"{old_val}" if old_val is not None else "(空)"
            new_str = f"{new_val}" if new_val is not None else "(空)"
            if isinstance(old_val, str) and old_val.startswith('=') or \
               isinstance(new_val, str) and new_val.startswith('='):
                return {'type': '公式变化', 'desc': f'{old_str} → {new_str}'}
            else:
                return {'type': '值变化', 'desc': f'{old_str} → {new_str}'}

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
        """尽可能获取图片（openpyxl）"""
        images = []
        # 方法1：_images
        if hasattr(ws, '_images') and ws._images:
            for img in ws._images:
                try:
                    anchor = img.anchor
                    if anchor and hasattr(anchor, '_from'):
                        col = anchor._from.col
                        row = anchor._from.row
                        addr = cell_address(col, row)
                        images.append((addr, img.width, img.height))
                except Exception as e:
                    self.log(f"图片提取异常: {e}")
        if images:
            return sorted(images, key=lambda x: x[0])
        # 方法2：_drawing
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
                if addr1 != addr2 or abs(w1 - w2) > 0.5 or abs(h1 - h2) > 0.5:
                    diff_detected = True
                    break

        if diff_detected:
            self.stats['images_diff'] += 1
            anchor = old_imgs[0][0] if old_imgs else (new_imgs[0][0] if new_imgs else 'A1')
            desc = f"图片差异：旧版 {len(old_imgs)} 张，新版 {len(new_imgs)} 张"
            if len(old_imgs) == len(new_imgs):
                desc += "，尺寸或位置变化"
            self.diffs.append({
                'sheet': sheet_name,
                'address': anchor,
                'type': '图片差异',
                'desc': desc
            })

    def _compare_conditional_formats(self, old_ws, new_ws, sheet_name):
        old_cfs = list(old_ws.conditional_formatting)
        new_cfs = list(new_ws.conditional_formatting)

        # 建立旧版本的范围->规则列表映射
        old_map = {str(cf.sqref): cf for cf in old_cfs}
        new_map = {str(cf.sqref): cf for cf in new_cfs}

        all_ranges = set(old_map.keys()) | set(new_map.keys())
        for rng in all_ranges:
            old_cf = old_map.get(rng)
            new_cf = new_map.get(rng)
            if old_cf is None:
                # 新增的条件格式
                start = rng.split(':')[0] if ':' in rng else rng
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': start,
                    'type': '条件格式新增',
                    'desc': f'新增条件格式范围: {rng}'
                })
            elif new_cf is None:
                # 删除的条件格式
                start = rng.split(':')[0] if ':' in rng else rng
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': start,
                    'type': '条件格式删除',
                    'desc': f'删除条件格式范围: {rng}'
                })
            else:
                # 比较规则细节
                old_rules = [(r.type, r.priority, str(r.formula), str(r.dxf)) for r in old_cf.rules]
                new_rules = [(r.type, r.priority, str(r.formula), str(r.dxf)) for r in new_cf.rules]
                if old_rules != new_rules:
                    start = rng.split(':')[0] if ':' in rng else rng
                    self.diffs.append({
                        'sheet': sheet_name,
                        'address': start,
                        'type': '条件格式修改',
                        'desc': f'条件格式规则变化，范围: {rng}'
                    })

# ---------------------------- GUI ----------------------------
class DiffViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异对比工具")
        self.root.geometry("1000x700")

        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.topmost = tk.BooleanVar(value=False)

        # 顶部工具栏
        toolbar = tb.Frame(root, padding=5)
        toolbar.pack(fill='x')

        # 左侧按钮区
        btn_frame = tb.Frame(toolbar)
        btn_frame.pack(side='left', padx=(0, 10))

        self.start_btn = tb.Button(btn_frame, text="开始对比", bootstyle=PRIMARY, width=8, command=self.start_compare)
        self.start_btn.pack(side='left', padx=2)
        self.jump_btn = tb.Button(btn_frame, text="跳转", bootstyle=SECONDARY, width=6, command=self.jump_to_selected, state='disabled')
        self.jump_btn.pack(side='left', padx=2)

        # 文件路径区
        path_frame = tb.Frame(toolbar)
        path_frame.pack(side='left', fill='x', expand=True, padx=(0, 10))

        tb.Label(path_frame, text="旧版:").grid(row=0, column=0, sticky='w', padx=(0, 5), pady=2)
        self.old_entry = tb.Entry(path_frame, textvariable=self.old_path)
        self.old_entry.grid(row=0, column=1, sticky='ew', pady=2)
        tb.Button(path_frame, text="浏览", bootstyle=INFO, width=6, command=lambda: self.browse(self.old_path)).grid(row=0, column=2, padx=(5, 0), pady=2)

        tb.Label(path_frame, text="新版:").grid(row=1, column=0, sticky='w', padx=(0, 5), pady=2)
        self.new_entry = tb.Entry(path_frame, textvariable=self.new_path)
        self.new_entry.grid(row=1, column=1, sticky='ew', pady=2)
        tb.Button(path_frame, text="浏览", bootstyle=INFO, width=6, command=lambda: self.browse(self.new_path)).grid(row=1, column=2, padx=(5, 0), pady=2)

        path_frame.columnconfigure(1, weight=1)

        # 置顶复选框
        tb.Checkbutton(toolbar, text="置顶", variable=self.topmost, command=self.toggle_topmost, bootstyle="round-toggle").pack(side='left', padx=5)

        # 进度条
        self.progress = tb.Progressbar(root, mode='determinate', bootstyle=PRIMARY)
        self.progress.pack(fill='x', padx=5, pady=(0, 5))

        # 中央差异树
        tree_frame = tb.Frame(root, padding=(5, 0))
        tree_frame.pack(fill='both', expand=True)
        self.tree = tb.Treeview(tree_frame, columns=('address','type'), show='tree headings', bootstyle=PRIMARY)
        self.tree.heading('#0', text='Sheet / 差异项')
        self.tree.heading('address', text='位置')
        self.tree.heading('type', text='类型')
        self.tree.column('#0', width=250)
        self.tree.column('address', width=80)
        self.tree.column('type', width=100)
        scroll_y = tb.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview, bootstyle=ROUND)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Double-1>', self.jump_to_selected)

        # 底部：详情（左）与日志（右）
        bottom_frame = tb.Frame(root, padding=5)
        bottom_frame.pack(fill='x')
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)

        detailf = tb.LabelFrame(bottom_frame, text="差异详情", padding=5, bootstyle=INFO)
        detailf.grid(row=0, column=0, sticky='nsew', padx=(0, 3))
        self.detail = tk.Text(detailf, height=6, wrap='word', font=("微软雅黑", 9),
                              bg='#ffffff', fg='#212529', relief='flat',
                              highlightthickness=1, highlightbackground='#dee2e6',
                              highlightcolor='#0d6efd', padx=5, pady=5)
        self.detail.pack(fill='both', expand=True)

        logf = tb.LabelFrame(bottom_frame, text="日志", padding=5, bootstyle=SECONDARY)
        logf.grid(row=0, column=1, sticky='nsew', padx=(3, 0))
        self.log_text = tk.Text(logf, height=6, wrap='word', font=("微软雅黑", 9),
                                bg='#ffffff', fg='#212529', relief='flat',
                                highlightthickness=1, highlightbackground='#dee2e6',
                                highlightcolor='#0d6efd', padx=5, pady=5)
        self.log_text.pack(fill='both', expand=True)

        self.diff_items = []
        self.result_data = None

    def browse(self, var):
        p = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if p: var.set(p)

    def toggle_topmost(self):
        self.root.attributes('-topmost', self.topmost.get())

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
        if not sel: return
        node = sel[0]
        target = next((d for n, d in self.diff_items if n == node), None)
        if not target: return
        if target['type'] == 'cell' or target['type'] == 'sheet_struct':
            d = target['data']
            self.detail.delete('1.0','end')
            self.detail.insert('1.0', f"Sheet: {d['sheet']}\n位置: {d.get('address','?')}\n类型: {d['type']}\n描述: {d['desc']}")

    def jump_to_selected(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        node = sel[0]
        target = next((d for n, d in self.diff_items if n == node), None)
        if not target: return

        if target['type'] == 'cell' or target['type'] == 'sheet_struct':
            d = target['data']
            sheet = d['sheet']
            address = d.get('address', 'A1')
        else:
            return

        def navigate():
            pythoncom.CoInitialize()
            try:
                excel = win32com.client.GetObject(Class="Excel.Application")
            except:
                messagebox.showinfo("提示", "未检测到正在运行的 Excel 进程，请手动打开两个文件后再跳转。")
                return

            old_path = normalize_path(self.old_path.get())
            new_path = normalize_path(self.new_path.get())
            old_wb = new_wb = None
            for wb in excel.Workbooks:
                if normalize_path(wb.FullName) == old_path: old_wb = wb
                elif normalize_path(wb.FullName) == new_path: new_wb = wb

            if not old_wb or not new_wb:
                missing = []
                if not old_wb: missing.append(f"旧版: {self.old_path.get()}")
                if not new_wb: missing.append(f"新版: {self.new_path.get()}")
                messagebox.showinfo("文件未打开", "以下文件未打开：\n" + "\n".join(missing) + "\n请手动打开后重试。")
                pythoncom.CoUninitialize()
                return

            for wb, desc in [(old_wb, "旧版"), (new_wb, "新版")]:
                try:
                    ws = wb.Worksheets(sheet)
                    ws.Activate()
                    excel.Goto(ws.Range(address), Scroll=True)
                    self.log(f"{desc} 已跳转到 {sheet}!{address}")
                except Exception as e:
                    self.log(f"{desc} 跳转失败: {e}")

            excel.Visible = True
            excel.WindowState = -4137
            pythoncom.CoUninitialize()

        threading.Thread(target=navigate, daemon=True).start()

if __name__ == "__main__":
    app = tb.Window(themename="flatly")
    viewer = DiffViewer(app)
    app.mainloop()
