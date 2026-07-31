"""
Excel 差异对比工具（最终增强版）- ttkbootstrap flatly 主题版
- 每个条件格式差异单独列出，精确定位首单元格
- 增强图片检测，openpyxl + COM 双重保障
- 界面优化：按钮左置，路径缩短，详情/日志对调，置顶开关
- 可靠富文本检测（zipfile + lxml 直接解析底层 XML）
- 修复：图片位置 0-based 偏移、富文本检测失效
"""
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import threading
import time
import os
import zipfile
import pythoncom
import win32com.client
from win32com.client import constants
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from lxml import etree

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

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

# ---------------------------- 富文本底层解析 ----------------------------
def _parse_rPr(rPr_elem):
    """从 rPr 元素解析字体属性字典"""
    font = {'name': '', 'size': '', 'bold': False, 'italic': False, 'underline': False, 'color': ''}
    if rPr_elem is None:
        return font
    rf = rPr_elem.find(f'{NS}rFont')
    if rf is not None:
        font['name'] = rf.get('val', '')
    sz = rPr_elem.find(f'{NS}sz')
    if sz is not None:
        font['size'] = sz.get('val', '')
    font['bold'] = rPr_elem.find(f'{NS}b') is not None
    font['italic'] = rPr_elem.find(f'{NS}i') is not None
    font['underline'] = rPr_elem.find(f'{NS}u') is not None
    color = rPr_elem.find(f'{NS}color')
    if color is not None:
        font['color'] = color.get('rgb', '') or color.get('theme', '')
    return font

def _extract_runs_from_si(si_elem):
    """从 sharedStrings 的 <si> 元素提取 run 列表，返回 [(text, font_dict)]"""
    runs = []
    # 普通文本 <t>
    t_elem = si_elem.find(f'{NS}t')
    if t_elem is not None:
        runs.append((t_elem.text or '', None))
        return runs
    # 富文本 <r>
    for r in si_elem.findall(f'{NS}r'):
        t_elem = r.find(f'{NS}t')
        text = t_elem.text if t_elem is not None else ''
        rPr = r.find(f'{NS}rPr')
        font = _parse_rPr(rPr)
        runs.append((text, font))
    return runs

def parse_rich_text_from_xlsx(xlsx_path):
    """
    直接从 xlsx 底层 XML 解析富文本信息。
    返回：{sheet_name: {cell_ref: [(text, font_dict_or_None), ...]}}
    只有确实存在富文本格式（非纯文本）的单元格才会被包含。
    """
    result = {}
    if not os.path.isfile(xlsx_path):
        return result

    try:
        zf = zipfile.ZipFile(xlsx_path)
    except:
        return result

    # 1. 解析 sharedStrings.xml，建立 索引 -> run列表 映射
    shared_runs = []
    if 'xl/sharedStrings.xml' in zf.namelist():
        ss_xml = zf.read('xl/sharedStrings.xml')
        ss_root = etree.fromstring(ss_xml)
        for si in ss_root.findall(f'{NS}si'):
            runs = _extract_runs_from_si(si)
            shared_runs.append(runs)

    # 2. 解析 workbook.xml + rels，正确映射 sheet 文件名与显示名
    sheet_name_map = {}  # sheet文件名(不含路径) -> sheet显示名
    R_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'

    # 先读 workbook.xml.rels，建立 rId -> target 路径映射
    rid_to_target = {}
    if 'xl/_rels/workbook.xml.rels' in zf.namelist():
        rels_xml = zf.read('xl/_rels/workbook.xml.rels')
        rels_root = etree.fromstring(rels_xml)
        for rel in rels_root.findall(f'{R_NS}Relationship'):
            rid = rel.get('Id', '')
            target = rel.get('Target', '')
            rid_to_target[rid] = target

    # 再读 workbook.xml，拿到 sheet 名和 rId 的对应
    if 'xl/workbook.xml' in zf.namelist():
        wb_xml = zf.read('xl/workbook.xml')
        wb_root = etree.fromstring(wb_xml)
        sheets_elem = wb_root.find(f'{NS}sheets')
        if sheets_elem is not None:
            R_NS_IN_WB = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            for sheet_elem in sheets_elem.findall(f'{NS}sheet'):
                name = sheet_elem.get('name', '')
                rid = sheet_elem.get(f'{R_NS_IN_WB}id', '')
                if rid and rid in rid_to_target:
                    target = rid_to_target[rid]
                    # target 可能是 worksheets/sheet1.xml 或 /xl/worksheets/sheet1.xml
                    filename = target.split('/')[-1]
                    if filename.endswith('.xml'):
                        sheet_name_map[filename] = name

    # 兜底：如果 rels 方式没拿到，按 sheet 序号猜
    if not sheet_name_map:
        sheet_files = sorted(
            [n for n in zf.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')],
            key=lambda x: int(x.split('sheet')[1].split('.xml')[0]) if 'sheet' in x else 0
        )
        for idx, f in enumerate(sheet_files, 1):
            sheet_name_map[f.split('/')[-1]] = f'Sheet{idx}'

    # 3. 逐个解析 worksheet XML
    for sheet_file, sheet_name in sheet_name_map.items():
        sheet_path = f'xl/worksheets/{sheet_file}'
        if sheet_path not in zf.namelist():
            continue
        try:
            ws_xml = zf.read(sheet_path)
        except:
            continue
        ws_root = etree.fromstring(ws_xml)
        sheet_data = {}
        for c in ws_root.findall(f'.//{NS}c'):
            ref = c.get('r', '')
            t = c.get('t', '')  # 类型: s=共享字符串, inlineStr=内嵌富文本, str=公式字符串...
            if t == 's':
                # 共享字符串
                v_elem = c.find(f'{NS}v')
                if v_elem is not None and v_elem.text is not None:
                    try:
                        idx = int(v_elem.text)
                        if 0 <= idx < len(shared_runs):
                            runs = shared_runs[idx]
                            if runs and any(f is not None for _, f in runs):
                                sheet_data[ref] = runs
                    except (ValueError, IndexError):
                        pass
            elif t == 'inlineStr':
                # 内嵌富文本
                is_elem = c.find(f'{NS}is')
                if is_elem is not None:
                    runs = _extract_runs_from_si(is_elem)
                    if runs and any(f is not None for _, f in runs):
                        sheet_data[ref] = runs
        if sheet_data:
            result[sheet_name] = sheet_data

    zf.close()
    return result

def compare_rich_text_runs(runs1, runs2):
    """比较两列富文本 run，返回差异描述或 None"""
    if (not runs1 and not runs2):
        return None
    if runs1 is None and runs2 is None:
        return None

    # 一边有富文本一边没有
    if (runs1 and not runs2) or (not runs1 and runs2):
        plain1 = ''.join(t for t, f in (runs1 or []))
        plain2 = ''.join(t for t, f in (runs2 or []))
        if plain1 != plain2:
            return f"内容(含富文本): {plain1} → {plain2}"
        return "单元格变为富文本格式"

    plain1 = ''.join(t for t, f in runs1)
    plain2 = ''.join(t for t, f in runs2)
    if plain1 != plain2:
        return f"内容(含富文本): {plain1} → {plain2}"

    if len(runs1) != len(runs2):
        return f"富文本段落数不同: {len(runs1)} → {len(runs2)}"

    changes = []
    for i, ((t1, f1), (t2, f2)) in enumerate(zip(runs1, runs2)):
        if f1 != f2:
            diff = []
            for k in ['name', 'size', 'bold', 'italic', 'underline', 'color']:
                v1 = f1.get(k) if f1 else ''
                v2 = f2.get(k) if f2 else ''
                if v1 != v2:
                    diff.append(f"{k}: {v1}→{v2}")
            if diff:
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
        self.old_rich = {}  # {sheet: {ref: runs}}
        self.new_rich = {}

    def run(self):
        self.progress(3, "解析富文本结构...")
        self.old_rich = parse_rich_text_from_xlsx(self.old_path)
        self.new_rich = parse_rich_text_from_xlsx(self.new_path)
        self.log(f"富文本解析完成：旧版 {sum(len(v) for v in self.old_rich.values())} 个富文本单元格，"
                 f"新版 {sum(len(v) for v in self.new_rich.values())} 个")

        self.progress(8, "加载旧版文件...")
        old_wb = load_workbook(self.old_path, data_only=False)
        self.progress(18, "加载新版文件...")
        new_wb = load_workbook(self.new_path, data_only=False)

        self._compare_sheets(old_wb, new_wb)
        common = set(old_wb.sheetnames) & set(new_wb.sheetnames)
        total = len(common)
        for idx, sheet_name in enumerate(common, 1):
            self.progress(25 + int(65 * idx / total), f"对比 {sheet_name}...")
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

        old_sheet_rich = self.old_rich.get(sheet_name, {})
        new_sheet_rich = self.new_rich.get(sheet_name, {})

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                old_cell = old_ws.cell(row, col)
                new_cell = new_ws.cell(row, col)
                ref = cell_address(col, row)
                diff = self._get_cell_diff(old_cell, new_cell, ref,
                                           old_sheet_rich.get(ref),
                                           new_sheet_rich.get(ref))
                if diff:
                    self.stats['diff_cells'] += 1
                    self.diffs.append({
                        'sheet': sheet_name,
                        'address': ref,
                        'type': diff['type'],
                        'desc': diff['desc']
                    })

        self._compare_merged_cells(old_ws, new_ws, sheet_name)
        self._compare_row_col_dimensions(old_ws, new_ws, sheet_name)
        self._compare_images(old_ws, new_ws, sheet_name)
        self._compare_conditional_formats(old_ws, new_ws, sheet_name)

    def _get_cell_diff(self, old_cell, new_cell, ref, old_runs, new_runs):
        # 优先检测富文本（底层XML方式，更可靠）
        if old_runs is not None or new_runs is not None:
            # 纯文本侧包装成无格式 run，确保内容比较正确
            r1 = old_runs if old_runs is not None else (
                [(str(old_cell.value), None)] if old_cell.value is not None else [])
            r2 = new_runs if new_runs is not None else (
                [(str(new_cell.value), None)] if new_cell.value is not None else [])
            rich_diff = compare_rich_text_runs(r1, r2)
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
                col_idx = column_index_from_string(col_letter)
                addr = cell_address(col_idx, 1)
                self.diffs.append({'sheet':sheet_name,'address':addr,'type':'列宽变化','desc':f'列宽({col_letter}): {ow} → {nw}'})
        for col_letter, cd in new_ws.column_dimensions.items():
            if col_letter not in old_ws.column_dimensions:
                nw = cd.width
                if nw is not None:
                    col_idx = column_index_from_string(col_letter)
                    addr = cell_address(col_idx, 1)
                    self.diffs.append({'sheet':sheet_name,'address':addr,'type':'列宽新设置','desc':f'列宽({col_letter}): {nw}'})

    def _get_images_from_ws(self, ws):
        """获取图片列表，返回 [(address, width, height)] —— address 使用 1-based 单元格坐标"""
        images = []
        # 方法1：_images
        if hasattr(ws, '_images') and ws._images:
            for img in ws._images:
                try:
                    anchor = img.anchor
                    if anchor and hasattr(anchor, '_from'):
                        col = anchor._from.col + 1  # 0-based → 1-based
                        row = anchor._from.row + 1
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
                    col = (anchor._from.col if hasattr(anchor, '_from') else 0) + 1
                    row = (anchor._from.row if hasattr(anchor, '_from') else 0) + 1
                    addr = cell_address(col, row)
                    images.append((addr, img.width, img.height))
        return sorted(images, key=lambda x: x[0])

    def _compare_images(self, old_ws, new_ws, sheet_name):
        old_imgs = self._get_images_from_ws(old_ws)
        new_imgs = self._get_images_from_ws(new_ws)
        if not old_imgs and not new_imgs:
            return

        old_set = {addr: (w, h) for addr, w, h in old_imgs}
        new_set = {addr: (w, h) for addr, w, h in new_imgs}

        old_addrs = set(old_set.keys())
        new_addrs = set(new_set.keys())

        added = new_addrs - old_addrs
        removed = old_addrs - new_addrs
        common = old_addrs & new_addrs

        changed = []
        for addr in common:
            w1, h1 = old_set[addr]
            w2, h2 = new_set[addr]
            # 使用相对误差 + 绝对误差组合判断尺寸变化
            w_rel = abs(w1 - w2) / max(w1, w2) if max(w1, w2) > 0 else 0
            h_rel = abs(h1 - h2) / max(h1, h2) if max(h1, h2) > 0 else 0
            w_abs = abs(w1 - w2)
            h_abs = abs(h1 - h2)
            if w_rel > 0.01 or h_rel > 0.01:  # 相对变化超过1%
                if w_abs > 1 or h_abs > 1:    # 且绝对变化超过1像素
                    changed.append((addr, w1, h1, w2, h2))

        diff_count = len(added) + len(removed) + len(changed)
        if diff_count > 0:
            self.stats['images_diff'] += diff_count
            # 每条差异单独列出，方便逐条跳转
            for addr in sorted(added):
                w, h = new_set[addr]
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': addr,
                    'type': '图片新增',
                    'desc': f'新增图片 ({w:.0f}x{h:.0f})'
                })
            for addr in sorted(removed):
                w, h = old_set[addr]
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': addr,
                    'type': '图片删除',
                    'desc': f'删除图片 ({w:.0f}x{h:.0f})'
                })
            for addr, w1, h1, w2, h2 in sorted(changed, key=lambda x: x[0]):
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': addr,
                    'type': '图片尺寸变化',
                    'desc': f'图片尺寸: {w1:.0f}x{h1:.0f} → {w2:.0f}x{h2:.0f}'
                })

    def _compare_conditional_formats(self, old_ws, new_ws, sheet_name):
        old_cfs = list(old_ws.conditional_formatting)
        new_cfs = list(new_ws.conditional_formatting)

        old_map = {str(cf.sqref): cf for cf in old_cfs}
        new_map = {str(cf.sqref): cf for cf in new_cfs}

        all_ranges = set(old_map.keys()) | set(new_map.keys())
        for rng in all_ranges:
            old_cf = old_map.get(rng)
            new_cf = new_map.get(rng)
            if old_cf is None:
                start = rng.split(':')[0] if ':' in rng else rng
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': start,
                    'type': '条件格式新增',
                    'desc': f'新增条件格式范围: {rng}'
                })
            elif new_cf is None:
                start = rng.split(':')[0] if ':' in rng else rng
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': start,
                    'type': '条件格式删除',
                    'desc': f'删除条件格式范围: {rng}'
                })
            else:
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

        # ========== 顶部工具栏（grid布局，按钮上下对齐） ==========
        toolbar = tb.Frame(root, padding=5)
        toolbar.pack(fill='x')
        toolbar.columnconfigure(3, weight=1)   # 路径区域整体可拉伸

        # 开始对比按钮（占2行，info样式）
        self.start_btn = tb.Button(toolbar, text="开始对比", bootstyle=INFO, width=8, command=self.start_compare)
        self.start_btn.grid(row=0, column=0, rowspan=2, sticky='nsew', padx=2, pady=1)

        # 跳转按钮（占2行）
        self.jump_btn = tb.Button(toolbar, text="跳转", bootstyle=SECONDARY, width=6, command=self.jump_to_selected, state='disabled')
        self.jump_btn.grid(row=0, column=1, rowspan=2, sticky='nsew', padx=2, pady=1)

        # 分隔
        tb.Separator(toolbar, orient='vertical').grid(row=0, column=2, rowspan=2, sticky='ns', padx=8)

        # 路径区域 frame（内部自己管理标签+输入框+浏览，输入框占满剩余空间）
        path_frame = tb.Frame(toolbar)
        path_frame.grid(row=0, column=3, rowspan=2, sticky='nsew', padx=(0, 10))
        path_frame.columnconfigure(1, weight=1)  # 输入框列可拉伸

        # 旧版行
        tb.Label(path_frame, text="旧版:").grid(row=0, column=0, sticky='w', padx=(0, 5), pady=2)
        self.old_entry = tb.Entry(path_frame, textvariable=self.old_path)
        self.old_entry.grid(row=0, column=1, sticky='ew', pady=2)
        tb.Button(path_frame, text="浏览", bootstyle=PRIMARY, width=6, command=lambda: self.browse(self.old_path)).grid(row=0, column=2, padx=(5, 0), pady=2)

        # 新版行
        tb.Label(path_frame, text="新版:").grid(row=1, column=0, sticky='w', padx=(0, 5), pady=2)
        self.new_entry = tb.Entry(path_frame, textvariable=self.new_path)
        self.new_entry.grid(row=1, column=1, sticky='ew', pady=2)
        tb.Button(path_frame, text="浏览", bootstyle=PRIMARY, width=6, command=lambda: self.browse(self.new_path)).grid(row=1, column=2, padx=(5, 0), pady=2)

        # 置顶开关（占2行，右侧）
        tb.Checkbutton(toolbar, text="置顶", variable=self.topmost, command=self.toggle_topmost, bootstyle="round-toggle").grid(
            row=0, column=4, rowspan=2, padx=(0, 5), pady=2, sticky='w')

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

        # ========== 底部：差异详情（固定宽） + 日志（可拉伸） ==========
        bottom_frame = tb.Frame(root, padding=5)
        bottom_frame.pack(fill='x')
        bottom_frame.columnconfigure(0, weight=0)   # 差异详情固定宽度
        bottom_frame.columnconfigure(1, weight=1)   # 日志随窗口拉伸

        # 差异详情（固定宽度，不随窗口缩放）
        detailf = tb.Labelframe(bottom_frame, text="差异详情", padding=5, bootstyle=INFO)
        detailf.grid(row=0, column=0, sticky='nsew', padx=(0, 3))
        detailf.grid_propagate(True)
        self.detail = tk.Text(detailf, width=42, height=6, wrap='word', font=("微软雅黑", 9),
                              bg='#ffffff', fg='#212529', relief='flat',
                              highlightthickness=1, highlightbackground='#dee2e6',
                              highlightcolor='#0d6efd', padx=5, pady=5)
        self.detail.pack(fill='both', expand=True)

        # 日志（可拉伸）
        logf = tb.Labelframe(bottom_frame, text="日志", padding=5, bootstyle=SECONDARY)
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
