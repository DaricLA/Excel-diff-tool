"""
Excel 差异对比工具（终极优化版）
- 界面：按钮与文件选择同排，差异树全屏中央，底部左侧日志、右侧详情各半
- 条件格式定位到其作用范围的首个单元格
- 富文本检测采用底层 XML 解析，确保高可靠性
- 跳转使用 Application.Goto，视觉效果更醒目
- 手动打开 Excel，不自动启动
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

# ---------------------------- 富文本检测模块 ----------------------------
def extract_rich_text_elements(cell_xml):
    """从单元格XML中提取所有<r>元素，返回列表，每个元素为 (text, font_dict)"""
    elements = []
    for r in cell_xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}r'):
        t = r.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
        text = t.text if t is not None else ""
        rPr = r.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}rPr')
        font = {}
        if rPr is not None:
            # 提取字体属性
            font['name'] = rPr.findtext('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}rFont', '')
            font['size'] = rPr.findtext('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sz', '')
            font['bold'] = rPr.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}b') is not None
            font['italic'] = rPr.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}i') is not None
            font['underline'] = rPr.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}u') is not None
            color_elem = rPr.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}color')
            if color_elem is not None:
                font['color'] = color_elem.get('rgb', 'None')
            else:
                font['color'] = 'None'
        elements.append((text, font))
    return elements

def compare_rich_text(cell1, cell2):
    """比较两个单元格的富文本，返回差异描述或None"""
    # 获取底层 XML
    xml1 = cell1._cell.xml if hasattr(cell1, '_cell') else None
    xml2 = cell2._cell.xml if hasattr(cell2, '_cell') else None
    if not xml1 or not xml2:
        return None

    elem1 = extract_rich_text_elements(etree.fromstring(xml1.encode()))
    elem2 = extract_rich_text_elements(etree.fromstring(xml2.encode()))

    if len(elem1) == 0 and len(elem2) == 0:
        return None  # 无富文本

    # 比较纯文本
    plain1 = ''.join(e[0] for e in elem1)
    plain2 = ''.join(e[0] for e in elem2)
    if plain1 != plain2:
        return f"内容(含富文本): {plain1} → {plain2}"

    # 比较段落格式
    if len(elem1) != len(elem2):
        return f"富文本段落数不同: {len(elem1)} → {len(elem2)}"
    changes = []
    for i, ((t1, f1), (t2, f2)) in enumerate(zip(elem1, elem2)):
        if f1 != f2:
            diff = []
            for key in ['name', 'size', 'bold', 'italic', 'underline', 'color']:
                if f1.get(key) != f2.get(key):
                    diff.append(f"{key}: {f1.get(key)}→{f2.get(key)}")
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
        # 1. 富文本优先检测
        rich_diff = compare_rich_text(old_cell, new_cell)
        if rich_diff:
            return {'type': '内容(含富文本)', 'desc': rich_diff}

        old_val = old_cell.value
        new_val = new_cell.value
        # 2. 值/公式变化（普通文本或公式）
        if old_val != new_val:
            old_str = f"{old_val}" if old_val is not None else "(空)"
            new_str = f"{new_val}" if new_val is not None else "(空)"
            if isinstance(old_val, str) and old_val.startswith('=') or \
               isinstance(new_val, str) and new_val.startswith('='):
                return {'type': '公式变化', 'desc': f'{old_str} → {new_str}'}
            else:
                return {'type': '值变化', 'desc': f'{old_str} → {new_str}'}

        # 3. 格式差异（整体字体、填充等）
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
            # 获取第一个条件格式范围的左上角单元格作为定位
            first_range = str(old_cfs[0].sqref) if old_cfs else str(new_cfs[0].sqref)
            start_cell = first_range.split(':')[0] if ':' in first_range else first_range
            self.diffs.append({
                'sheet': sheet_name,
                'address': start_cell,
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

        # 顶部控制栏：文件选择 + 按钮
        top = ttk.Frame(root, padding=5)
        top.grid(row=0, column=0, sticky='ew')
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="旧版:", font=("微软雅黑", 9)).grid(row=0,column=0,sticky='w')
        self.old_entry = ttk.Entry(top, textvariable=self.old_path, width=35)
        self.old_entry.grid(row=0,column=1,padx=5,sticky='ew')
        ttk.Button(top, text="浏览", command=lambda: self.browse(self.old_path)).grid(row=0,column=2)

        ttk.Label(top, text="新版:", font=("微软雅黑", 9)).grid(row=1,column=0,sticky='w',pady=3)
        self.new_entry = ttk.Entry(top, textvariable=self.new_path, width=35)
        self.new_entry.grid(row=1,column=1,padx=5,sticky='ew')
        ttk.Button(top, text="浏览", command=lambda: self.browse(self.new_path)).grid(row=1,column=2)

        btn_frame = ttk.Frame(top)
        btn_frame.grid(row=0,column=3,rowspan=2,padx=10)
        self.start_btn = tk.Button(btn_frame, text="开始对比", font=("微软雅黑", 10, "bold"),
                                   bg="#0078D7", fg="white", width=8, command=self.start_compare)
        self.start_btn.pack(side='left',padx=2)
        self.jump_btn = tk.Button(btn_frame, text="跳转", font=("微软雅黑", 10),
                                  command=self.jump_to_selected, state='disabled')
        self.jump_btn.pack(side='left',padx=2)

        self.progress = ttk.Progressbar(root, mode='determinate')
        self.progress.grid(row=1, column=0, sticky='ew', padx=5)

        # 中央：差异树
        self.tree = ttk.Treeview(root, columns=('address','type'), show='tree headings')
        self.tree.heading('#0',text='Sheet / 差异项'); self.tree.heading('address',text='位置'); self.tree.heading('type',text='类型')
        self.tree.column('#0',width=200); self.tree.column('address',width=80); self.tree.column('type',width=120)
        scroll_tree = ttk.Scrollbar(root, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_tree.set)
        self.tree.grid(row=2, column=0, sticky='nsew', padx=5, pady=5)
        scroll_tree.grid(row=2, column=1, sticky='ns')
        root.grid_rowconfigure(2, weight=1)   # 让树占据剩余高度

        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Double-1>', self.jump_to_selected)

        # 底部：左右分栏（日志 / 详情）
        bottom_frame = ttk.Frame(root)
        bottom_frame.grid(row=3, column=0, sticky='ew', padx=5, pady=(0,5))
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)

        logf = ttk.LabelFrame(bottom_frame, text="日志", padding=3)
        logf.grid(row=0, column=0, sticky='nsew', padx=(0,3))
        self.log_text = tk.Text(logf, height=6, wrap='word', font=("微软雅黑", 9))
        self.log_text.pack(fill='both',expand=True)

        detailf = ttk.LabelFrame(bottom_frame, text="差异详情", padding=3)
        detailf.grid(row=0, column=1, sticky='nsew')
        self.detail = tk.Text(detailf, height=6, wrap='word', font=("微软雅黑", 9))
        self.detail.pack(fill='both',expand=True)

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
        if not sel: return
        node = sel[0]
        target = next((d for n, d in self.diff_items if n == node), None)
        if not target: return
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
        target = next((d for n, d in self.diff_items if n == node), None)
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
            self.root.after(100, lambda: self.root.lower())
            pythoncom.CoUninitialize()

        threading.Thread(target=navigate, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = DiffViewer(root)
    root.mainloop()
