import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import threading
import time
import os
import sys
import json
import zipfile
import re
import copy
import pythoncom
import win32com.client
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from lxml import etree

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
PROGRAM_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_CHECK_OPTIONS = {
    'value': True, 'formula': True, 'rich_text': True, 'font': True,
    'fill': True, 'border': True, 'alignment': True, 'number_format': True,
    'merged_cells': True, 'row_height': True, 'col_width': True,
    'images': True, 'conditional_format': True
}

CHECK_OPTION_LABELS = {
    'value': '值变化', 'formula': '公式变化', 'rich_text': '富文本', 'font': '字体',
    'fill': '填充/背景色', 'border': '边框', 'alignment': '对齐方式',
    'number_format': '数字格式', 'merged_cells': '合并单元格', 'row_height': '行高',
    'col_width': '列宽', 'images': '图片', 'conditional_format': '条件格式'
}

CHECK_OPTION_GROUPS = [
    ("内容检测", ['value', 'formula', 'rich_text']),
    ("格式检测", ['font', 'fill', 'border', 'alignment', 'number_format']),
    ("结构检测", ['merged_cells', 'row_height', 'col_width', 'images', 'conditional_format'])
]

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

def normalize_color_for_compare(color):
    if color is None:
        return None
    try:
        if hasattr(color, 'type') and color.type == 'auto':
            return None
        if hasattr(color, 'theme') and color.theme is not None:
            return ('theme', color.theme, round(color.tint if color.tint is not None else 0.0, 4))
        if hasattr(color, 'indexed') and color.indexed is not None:
            return ('indexed', color.indexed)
        if hasattr(color, 'rgb') and color.rgb is not None:
            rgb = color.rgb
            if isinstance(rgb, str):
                if len(rgb) == 8 and rgb[:2] in ('00', 'FF'):
                    return ('rgb', rgb[2:].upper())
                return ('rgb', rgb.upper())
            return ('rgb', str(rgb).upper())
    except:
        pass
    return None

def get_sheet_names_fast(file_path):
    if not os.path.isfile(file_path):
        return []
    try:
        zf = zipfile.ZipFile(file_path)
        if 'xl/workbook.xml' in zf.namelist():
            root = etree.fromstring(zf.read('xl/workbook.xml'))
            sheets = [s.get('name') for s in root.findall(f'{NS}sheet')]
            zf.close()
            return sheets
    except:
        pass
    return []

def _parse_rPr(rPr):
    font = {'name': '', 'size': '', 'bold': False, 'italic': False, 'underline': False, 'color': ''}
    if rPr is None:
        return font
    rf = rPr.find(f'{NS}rFont')
    if rf is not None:
        font['name'] = rf.get('val', '')
    sz = rPr.find(f'{NS}sz')
    if sz is not None:
        font['size'] = sz.get('val', '')
    font['bold'] = rPr.find(f'{NS}b') is not None
    font['italic'] = rPr.find(f'{NS}i') is not None
    font['underline'] = rPr.find(f'{NS}u') is not None
    color = rPr.find(f'{NS}color')
    if color is not None:
        font['color'] = color.get('rgb', '') or color.get('theme', '')
    return font

def _extract_runs_from_si(si):
    runs = []
    t = si.find(f'{NS}t')
    if t is not None:
        return [(t.text or '', None)]
    for r in si.findall(f'{NS}r'):
        t = r.find(f'{NS}t')
        runs.append((t.text if t is not None else '', _parse_rPr(r.find(f'{NS}rPr'))))
    return runs

def parse_rich_text_from_xlsx(path):
    result = {}
    if not os.path.isfile(path):
        return result
    try:
        zf = zipfile.ZipFile(path)
    except:
        return result
    shared = []
    if 'xl/sharedStrings.xml' in zf.namelist():
        root = etree.fromstring(zf.read('xl/sharedStrings.xml'))
        for si in root.findall(f'{NS}si'):
            shared.append(_extract_runs_from_si(si))
    sheet_map = {}
    R_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    rid_to_target = {}
    if 'xl/_rels/workbook.xml.rels' in zf.namelist():
        root = etree.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
        for rel in root.findall(f'{R_NS}Relationship'):
            rid_to_target[rel.get('Id', '')] = rel.get('Target', '')
    if 'xl/workbook.xml' in zf.namelist():
        wb = etree.fromstring(zf.read('xl/workbook.xml'))
        sheets_elem = wb.find(f'{NS}sheets')
        if sheets_elem is not None:
            R_NS_IN_WB = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            for s in sheets_elem.findall(f'{NS}sheet'):
                name = s.get('name', '')
                rid = s.get(f'{R_NS_IN_WB}id', '')
                target = rid_to_target.get(rid, '')
                if target.endswith('.xml'):
                    sheet_map[target.split('/')[-1]] = name
    if not sheet_map:
        sheet_files = sorted(
            [n for n in zf.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')],
            key=lambda x: int(x.split('sheet')[1].split('.xml')[0])
        )
        for i, f in enumerate(sheet_files, 1):
            sheet_map[f.split('/')[-1]] = f'Sheet{i}'
    for filename, sheet_name in sheet_map.items():
        if f'xl/worksheets/{filename}' not in zf.namelist():
            continue
        ws = etree.fromstring(zf.read(f'xl/worksheets/{filename}'))
        data = {}
        for c in ws.findall(f'.//{NS}c'):
            ref = c.get('r', '')
            t = c.get('t', '')
            if t == 's':
                v = c.find(f'{NS}v')
                if v is not None and v.text:
                    idx = int(v.text)
                    if 0 <= idx < len(shared) and any(f is not None for _, f in shared[idx]):
                        data[ref] = shared[idx]
            elif t == 'inlineStr':
                is_ = c.find(f'{NS}is')
                if is_ is not None:
                    runs = _extract_runs_from_si(is_)
                    if runs and any(f is not None for _, f in runs):
                        data[ref] = runs
        if data:
            result[sheet_name] = data
    zf.close()
    return result

def compare_rich_text_runs(r1, r2):
    if (not r1 and not r2):
        return None
    if r1 is None and r2 is None:
        return None
    if (r1 and not r2) or (not r1 and r2):
        p1 = ''.join(t for t, f in (r1 or []))
        p2 = ''.join(t for t, f in (r2 or []))
        if p1 != p2:
            return f"内容(含富文本): {p1} → {p2}"
        return "单元格变为富文本格式"
    p1 = ''.join(t for t, f in r1)
    p2 = ''.join(t for t, f in r2)
    if p1 != p2:
        return f"内容(含富文本): {p1} → {p2}"
    if len(r1) != len(r2):
        return f"富文本段落数不同: {len(r1)} → {len(r2)}"
    changes = []
    for i, ((t1, f1), (t2, f2)) in enumerate(zip(r1, r2)):
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

# ======================== 语义定位引擎 ========================
class DataLocator:
    def __init__(self, rules_file=None):
        self.rules = []
        self.rules_file = rules_file
        if rules_file and os.path.isfile(rules_file):
            self.load_rules(rules_file)

    def load_rules(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.rules = data.get('rules', [])
            self.rules_file = filepath
            return True
        except:
            return False

    def locate_all(self, workbook):
        results = {}
        for rule in self.rules:
            name = rule.get('name', 'unnamed')
            try:
                results[name] = self.locate(workbook, rule)
            except Exception as e:
                results[name] = {'error': str(e)}
        return results

    def locate(self, workbook, rule):
        sheet_name = rule.get('sheet', '')
        if sheet_name not in workbook.sheetnames:
            return {'error': f'Sheet "{sheet_name}" not found'}
        ws = workbook[sheet_name]
        mode = rule.get('mode', 'offset')
        if mode == 'offset':
            return self._locate_offset(ws, rule)
        elif mode == 'collect':
            return self._locate_collect(ws, rule)
        elif mode == 'intersection':
            return self._locate_intersection(ws, rule)
        elif mode == 'range':
            return self._locate_range(ws, rule)
        return {'error': f'Unknown mode: {mode}'}

    def _find_anchor(self, ws, anchor_cfg):
        text = anchor_cfg.get('text', '').strip()
        search_in = anchor_cfg.get('search_in', 'all')
        if search_in == 'first_row':
            for col in range(1, (ws.max_column or 1) + 1):
                val = ws.cell(1, col).value
                if val is not None and text in str(val).strip():
                    return (1, col)
        elif search_in == 'first_col':
            for row in range(1, (ws.max_row or 1) + 1):
                val = ws.cell(row, 1).value
                if val is not None and text in str(val).strip():
                    return (row, 1)
        else:
            for row in range(1, (ws.max_row or 1) + 1):
                for col in range(1, (ws.max_column or 1) + 1):
                    val = ws.cell(row, col).value
                    if val is not None and text in str(val).strip():
                        return (row, col)
        return None

    def _locate_offset(self, ws, rule):
        anchor_cfg = rule.get('anchor', {})
        target = rule.get('target', {})
        row_offset = target.get('row_offset', 0)
        col_offset = target.get('col_offset', 0)
        anchor = self._find_anchor(ws, anchor_cfg)
        if not anchor:
            return {'error': f'Anchor "{anchor_cfg.get("text")}" not found'}
        target_row = anchor[0] + row_offset
        target_col = anchor[1] + col_offset
        if target_row < 1 or target_col < 1:
            return {'error': 'Target out of range'}
        cell = ws.cell(target_row, target_col)
        return {
            'address': cell_address(target_col, target_row),
            'value': cell.value,
            'is_formula': isinstance(cell.value, str) and cell.value.startswith('=')
        }

    def _locate_collect(self, ws, rule):
        anchor_cfg = rule.get('anchor', {})
        collect_cfg = rule.get('collect', {})
        direction = collect_cfg.get('direction', 'down')
        start_offset = collect_cfg.get('start_offset', 1)
        max_count = collect_cfg.get('max_count', 1000)
        anchor = self._find_anchor(ws, anchor_cfg)
        if not anchor:
            return {'error': f'Anchor "{anchor_cfg.get("text")}" not found'}
        data = []
        if direction == 'down':
            start_row = anchor[0] + start_offset
            col = anchor[1]
            for row in range(start_row, min(start_row + max_count, (ws.max_row or 1) + 1)):
                val = ws.cell(row, col).value
                if val is None:
                    break
                data.append({'row': row, 'value': val})
        elif direction == 'right':
            start_col = anchor[1] + start_offset
            row = anchor[0]
            for col in range(start_col, min(start_col + max_count, (ws.max_column or 1) + 1)):
                val = ws.cell(row, col).value
                if val is None:
                    break
                data.append({'col': col, 'value': val})
        return {
            'anchor_address': cell_address(anchor[1], anchor[0]),
            'direction': direction,
            'count': len(data),
            'values': data
        }

    def _locate_intersection(self, ws, rule):
        row_anchor_cfg = rule.get('row_anchor', {})
        col_anchor_cfg = rule.get('col_anchor', {})
        row_anchor = self._find_anchor(ws, {**row_anchor_cfg, 'search_in': row_anchor_cfg.get('search_in', 'all')})
        col_anchor = self._find_anchor(ws, {**col_anchor_cfg, 'search_in': col_anchor_cfg.get('search_in', 'all')})
        if not row_anchor:
            return {'error': f'Row anchor "{row_anchor_cfg.get("text")}" not found'}
        if not col_anchor:
            return {'error': f'Col anchor "{col_anchor_cfg.get("text")}" not found'}
        target_row = row_anchor[0]
        target_col = col_anchor[1]
        cell = ws.cell(target_row, target_col)
        return {
            'row_address': cell_address(1, target_row),
            'col_address': cell_address(target_col, 1),
            'address': cell_address(target_col, target_row),
            'value': cell.value,
            'is_formula': isinstance(cell.value, str) and cell.value.startswith('=')
        }

    def _locate_range(self, ws, rule):
        anchor_cfg = rule.get('anchor', {})
        target = rule.get('target', {})
        row_offset = target.get('row_offset', 0)
        col_offset = target.get('col_offset', 0)
        row_count = target.get('row_count', 1)
        col_count = target.get('col_count', 1)
        exclude = target.get('exclude', [])

        anchor = self._find_anchor(ws, anchor_cfg)
        if not anchor:
            return {'error': f'Anchor "{anchor_cfg.get("text")}" not found'}
        start_row = anchor[0] + row_offset
        start_col = anchor[1] + col_offset
        if start_row < 1 or start_col < 1:
            return {'error': 'Target out of range'}

        exclude_set = set()
        for ex in exclude:
            if isinstance(ex, str):
                m = re.match(r'^([A-Z]+)(\d+)$', ex)
                if m:
                    exclude_set.add(f"{m.group(1)}{m.group(2)}")
            elif isinstance(ex, list) and len(ex) == 2:
                exclude_set.add(cell_address(start_col + ex[1], start_row + ex[0]))

        addresses = []
        values = []
        for r in range(start_row, start_row + row_count):
            if r > (ws.max_row or 1):
                break
            for c in range(start_col, start_col + col_count):
                if c > (ws.max_column or 1):
                    break
                addr = cell_address(c, r)
                if addr in exclude_set:
                    continue
                cell = ws.cell(r, c)
                addresses.append(addr)
                values.append(cell.value)
        return {
            'address': addresses[0] if addresses else None,
            'addresses': addresses,
            'values': values,
            'range_count': len(addresses)
        }

# ======================== 插件框架 ========================
class CheckPlugin:
    name = ""
    description = ""
    def __init__(self, config=None):
        self.config = config or {}
    def check(self, old_data, new_data, context=None):
        raise NotImplementedError

class MeanDeviationPlugin(CheckPlugin):
    name = "mean_deviation"
    description = "均值偏差"
    def check(self, old_data, new_data, context=None):
        results = []
        threshold = self.config.get('threshold', 0.05)
        old_val = old_data.get('value') if isinstance(old_data, dict) else old_data
        new_val = new_data.get('value') if isinstance(new_data, dict) else new_data
        if old_val is None or new_val is None:
            return results
        try:
            old_num = float(old_val)
            new_num = float(new_val)
            if old_num != 0:
                dev = abs(new_num - old_num) / abs(old_num)
                if dev > threshold:
                    results.append({
                        'type': '均值偏差告警',
                        'desc': f'偏差 {dev:.2%} (阈值 {threshold:.0%}): {old_num:.4f} → {new_num:.4f}',
                        'severity': 'warning'
                    })
        except:
            pass
        return results

class ParamLockPlugin(CheckPlugin):
    name = "param_lock"
    description = "参数锁定"
    def check(self, old_data, new_data, context=None):
        results = []
        old_val = old_data.get('value') if isinstance(old_data, dict) else old_data
        new_val = new_data.get('value') if isinstance(new_data, dict) else new_data
        if old_val is not None and new_val is not None and old_val != new_val:
            results.append({
                'type': '参数修改告警',
                'desc': f'参数被修改: {old_val} → {new_val}',
                'severity': 'error'
            })
        return results

class RangeCheckPlugin(CheckPlugin):
    name = "range_check"
    description = "范围检查"
    def check(self, old_data, new_data, context=None):
        results = []
        lsl = self.config.get('lsl')
        usl = self.config.get('usl')
        new_val = new_data.get('value') if isinstance(new_data, dict) else new_data
        if new_val is None:
            return results
        try:
            num = float(new_val)
            if lsl is not None and num < float(lsl):
                results.append({
                    'type': '低于下限',
                    'desc': f'值 {num:.4f} < LSL {lsl}',
                    'severity': 'error'
                })
            if usl is not None and num > float(usl):
                results.append({
                    'type': '超出上限',
                    'desc': f'值 {num:.4f} > USL {usl}',
                    'severity': 'error'
                })
        except:
            pass
        return results

PLUGIN_REGISTRY = {
    'mean_deviation': MeanDeviationPlugin,
    'param_lock': ParamLockPlugin,
    'range_check': RangeCheckPlugin,
}

class PluginManager:
    def __init__(self, config_file=None):
        self.plugins = []
        self.locator = None
        if config_file and os.path.isfile(config_file):
            self.load_config(config_file)

    def load_config(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if 'locator_rules' in data:
                self.locator = DataLocator()
                self.locator.rules = data['locator_rules']
            for plugin_cfg in data.get('plugins', []):
                plugin_name = plugin_cfg.get('type', '')
                if plugin_name in PLUGIN_REGISTRY:
                    plugin = PLUGIN_REGISTRY[plugin_name](plugin_cfg.get('config', {}))
                    plugin.rule_name = plugin_cfg.get('rule_name', '')
                    plugin.description = plugin_cfg.get('description', '')
                    self.plugins.append(plugin)
            return True
        except:
            return False

    def run_checks(self, old_wb, new_wb, log_callback=None):
        if not self.locator or not self.plugins:
            return []
        results = []
        old_data = self.locator.locate_all(old_wb)
        new_data = self.locator.locate_all(new_wb)
        for plugin in self.plugins:
            rule_name = getattr(plugin, 'rule_name', '')
            if not rule_name:
                continue
            old_val = old_data.get(rule_name)
            new_val = new_data.get(rule_name)
            if old_val is None and new_val is None:
                continue
            try:
                diffs = plugin.check(old_val, new_val)
                for diff in diffs:
                    diff['rule_name'] = rule_name
                    diff['plugin'] = plugin.name
                    results.append(diff)
                    if log_callback:
                        log_callback(f" [{plugin.name}] {rule_name}: {diff['desc']}")
            except Exception as e:
                if log_callback:
                    log_callback(f" [{plugin.name}] {rule_name} 检查失败: {e}")
        return results

# ======================== 数据结构 ========================
class CheckItemConfig:
    def __init__(self, check_type="value", enabled=True, expect="same", options=None):
        self.check_type = check_type
        self.enabled = enabled
        self.expect = expect
        self.options = options or {}
    def to_dict(self):
        return {
            "check_type": self.check_type,
            "enabled": self.enabled,
            "expect": self.expect,
            "options": self.options
        }
    @classmethod
    def from_dict(cls, data):
        return cls(
            check_type=data.get("check_type", "value"),
            enabled=data.get("enabled", True),
            expect=data.get("expect", "same"),
            options=data.get("options", {})
        )

class CheckRule:
    def __init__(self, rule_name="", data_source=None, checks=None):
        self.rule_name = rule_name
        self.data_source = data_source or {}
        self.checks = checks or []
    def to_dict(self):
        return {
            "rule_name": self.rule_name,
            "data_source": self.data_source,
            "checks": [c.to_dict() for c in self.checks]
        }
    @classmethod
    def from_dict(cls, data):
        return cls(
            rule_name=data.get("rule_name", ""),
            data_source=data.get("data_source", {}),
            checks=[CheckItemConfig.from_dict(cd) for cd in data.get("checks", [])]
        )

class CheckProject:
    def __init__(self, project_name="", description="", version="1.0", rules=None):
        self.project_name = project_name
        self.description = description
        self.version = version
        self.rules = rules or []
    def to_dict(self):
        return {
            "project_name": self.project_name,
            "description": self.description,
            "version": self.version,
            "rules": [r.to_dict() for r in self.rules]
        }
    @classmethod
    def from_dict(cls, data):
        return cls(
            project_name=data.get("project_name", ""),
            description=data.get("description", ""),
            version=data.get("version", "1.0"),
            rules=[CheckRule.from_dict(rd) for rd in data.get("rules", [])]
        )

# ======================== 对比引擎 ========================
class OpenpyxlComparer:
    def __init__(self, old_path, new_path, log_callback=None, progress_callback=None,
                 check_options=None, plugin_manager=None, progress_mode_fn=None,
                 check_project=None, stop_event=None, mode='diff'):
        self.old_path = old_path
        self.new_path = new_path
        self.log = log_callback if log_callback else print
        self.progress = progress_callback if progress_callback else lambda v, s: None
        self.progress_mode = progress_mode_fn if progress_mode_fn else lambda m: None
        self.stop_event = stop_event if stop_event else threading.Event()
        self.mode = mode
        self.diffs = []
        self.sheet_diffs = []
        self.stats = {'total_cells': 0, 'diff_cells': 0, 'added_sheets': [], 'removed_sheets': [], 'images_diff': 0}
        self.old_rich = {}
        self.new_rich = {}
        self.check_options = dict(DEFAULT_CHECK_OPTIONS)
        self.check_options.update(check_options if check_options else {})
        self.plugin_manager = plugin_manager
        self.check_project = check_project
        self._log_buffer = []
        self._last_gui_update = 0

    def _flush_log(self, force=False):
        now = time.time()
        if not force and (now - self._last_gui_update) < 0.15:
            return
        if self._log_buffer:
            ts = time.strftime('%H:%M:%S')
            msg = '\n'.join(f"{ts} {m}" for m in self._log_buffer)
            self._log_buffer.clear()
            try:
                self.log(msg)
            except:
                pass
            self._last_gui_update = now

    def _buf_log(self, msg):
        self._log_buffer.append(msg)

    def run(self):
        start_time = time.time()
        old_wb, new_wb = self._load_workbooks()
        if not old_wb or not new_wb:
            return False
        self._run_diff_mode(old_wb, new_wb)
        self.progress(95, "生成报告...")
        self._flush_log(force=True)
        total_time = time.time() - start_time
        self.progress(100, "对比完成")
        self._buf_log(f"总耗时: {total_time:.1f}s | 差异: {self.stats['diff_cells']} 处单元格, {len(self.sheet_diffs)} 处Sheet")
        self._flush_log(force=True)
        return True

    def _load_workbooks(self):
        try:
            # 启动心跳线程刷新加载状态
            self._loading_msg = "正在加载旧版文件..."
            self.progress_mode('indeterminate')
            self.progress(5, "正在加载旧版文件...")
            self._flush_log(force=True)

            # 心跳线程（覆盖最后一行）
            heartbeat_stop = threading.Event()
            def heartbeat():
                start_time = time.time()
                while not heartbeat_stop.is_set():
                    time.sleep(1)
                    if heartbeat_stop.is_set():
                        break
                    msg = f"{self._loading_msg} 已耗时 {int(time.time()-start_time)}s"
                    self._buf_log(msg)
                    self._flush_log(force=True)
                    # 覆盖最后一行：删除末尾后重新插入
                    # 通过log机制实现（后续优化）
            hb_thread = threading.Thread(target=heartbeat, daemon=True)
            hb_thread.start()

            old_wb = load_workbook(self.old_path, data_only=False)
            self._loading_msg = "旧版加载完成"
            self._buf_log(f"旧版加载完成: {len(old_wb.sheetnames)} 个sheet")
            self._flush_log(force=True)

            self._loading_msg = "正在加载新版文件..."
            self.progress(15, "正在加载新版文件...")
            self._flush_log(force=True)
            new_wb = load_workbook(self.new_path, data_only=False)
            self._loading_msg = "新版加载完成"
            self._buf_log(f"新版加载完成: {len(new_wb.sheetnames)} 个sheet")
            self._flush_log(force=True)

            self._loading_msg = "正在解析富文本..."
            self.progress(20, "解析富文本...")
            self._flush_log(force=True)
            self.old_rich = parse_rich_text_from_xlsx(self.old_path)
            self.new_rich = parse_rich_text_from_xlsx(self.new_path)
            self._buf_log(f"富文本解析完成：旧版 {sum(len(v) for v in self.old_rich.values())} 个，新版 {sum(len(v) for v in self.new_rich.values())} 个")
            self._flush_log(force=True)

            heartbeat_stop.set()
            self.progress_mode('determinate')
            return old_wb, new_wb
        except Exception as e:
            heartbeat_stop.set()
            self.progress_mode('determinate')
            self._buf_log(f"加载工作簿失败: {e}")
            return None, None

    def _run_diff_mode(self, old_wb, new_wb):
        self._compare_sheets(old_wb, new_wb)
        total = len(old_wb.sheetnames)
        start_time = time.time()
        for idx, sheet_name in enumerate(old_wb.sheetnames, 1):
            if self.stop_event.is_set():
                self._buf_log(f"用户请求停止，已跳过剩余 {total - idx + 1} 个sheet")
                self._flush_log(force=True)
                raise KeyboardInterrupt
            pct = 25 + int(55 * idx / total)
            self.progress(pct, f"对比 {sheet_name}... ({idx}/{total})")
            self._flush_log(force=True)
            if sheet_name in new_wb.sheetnames:
                self._compare_worksheet(old_wb[sheet_name], new_wb[sheet_name], sheet_name)
            elapsed = time.time() - start_time
            self._buf_log(f"已完成 {sheet_name} ({idx}/{total})，累计耗时 {elapsed:.0f}s")
            self._flush_log(force=True)
        if self.plugin_manager and self.plugin_manager.plugins:
            self.progress(85, "执行数据检查插件...")
            self._buf_log(f"执行 {len(self.plugin_manager.plugins)} 个检查插件...")
            self._flush_log(force=True)
            plugin_results = self.plugin_manager.run_checks(old_wb, new_wb, self._buf_log)
            for diff in plugin_results:
                self.diffs.append({
                    'sheet': '🔍 数据检查',
                    'address': diff.get('rule_name', ''),
                    'type': diff['type'],
                    'desc': diff['desc']
                })
        if self.check_project:
            self.progress(90, "执行进阶规则过滤...")
            self._apply_rule_filter(self.diffs, old_wb, new_wb)

    def _apply_rule_filter(self, diffs, old_wb, new_wb):
        diff_type_map = {
            '内容变化': 'value',
            '公式变化': 'formula',
            '字体变化': 'font',
            '填充变化': 'fill',
            '边框变化': 'border',
            '对齐变化': 'alignment',
            '数字格式变化': 'number_format',
            '合并新增': 'merged_cells',
            '合并删除': 'merged_cells',
            '行高变化': 'row_height',
            '列宽变化': 'col_width',
            '图片新增': 'images',
            '图片删除': 'images',
            '图片尺寸变化': 'images',
            '条件格式新增': 'conditional_format',
            '条件格式删除': 'conditional_format',
            '条件格式修改': 'conditional_format',
            '富文本变化': 'rich_text',
            '单元格新增': 'value',
            '单元格删除': 'value'
        }
        rule_addr_map = {}
        for rule in self.check_project.rules:
            ds = rule.data_source
            sheet = ds.get('sheet', '')
            if sheet not in old_wb.sheetnames or sheet not in new_wb.sheetnames:
                continue
            locator = DataLocator()
            locator.rules = [ds]
            old_data = locator.locate_all(old_wb).get(ds.get('name', ''))
            new_data = locator.locate_all(new_wb).get(ds.get('name', ''))
            if not old_data or not new_data:
                continue
            addresses = old_data.get('addresses') or [old_data.get('address')] if isinstance(old_data, dict) else None
            if addresses:
                for addr in addresses:
                    rule_addr_map.setdefault((sheet, addr), []).append(rule)

        for d in diffs:
            if d['sheet'] == '🔍 数据检查':
                continue
            check_type = diff_type_map.get(d['type'])
            if not check_type:
                continue
            key = (d['sheet'], d['address'])
            if key not in rule_addr_map:
                continue
            for rule in rule_addr_map[key]:
                ds = rule.data_source
                old_ws = old_wb[d['sheet']]
                new_ws = new_wb[d['sheet']]
                col_str = ''.join(ch for ch in d['address'] if ch.isalpha())
                row_str = ''.join(ch for ch in d['address'] if ch.isdigit())
                if not col_str or not row_str:
                    continue
                col = column_index_from_string(col_str)
                row = int(row_str)
                old_cell = old_ws.cell(row=row, column=col)
                new_cell = new_ws.cell(row=row, column=col)
                for check in rule.checks:
                    if check.enabled and check.check_type == check_type:
                        diff = self._compare_by_check_type(check_type, old_cell, new_cell, check.options, old_ws, new_ws, d['address'], d['sheet'])
                        if check.expect == 'same' and diff is None:
                            d['rule_pass'] = True
                            d['rule_name'] = rule.rule_name
                            break
                        elif check.expect == 'different' and diff is not None:
                            d['rule_pass'] = True
                            d['rule_name'] = rule.rule_name
                            break
                if d.get('rule_pass'):
                    break

    def _compare_worksheet(self, old_ws, new_ws, sheet_name):
        opts = self.check_options
        _, _, old_real_max_row, old_real_max_col = self._real_data_range(old_ws)
        _, _, new_real_max_row, new_real_max_col = self._real_data_range(new_ws)
        max_row = max(old_real_max_row, new_real_max_row)
        max_col = max(old_real_max_col, new_real_max_col)
        max_row = min(max_row, 500000)
        max_col = min(max_col, 500)
        old_max_row = old_real_max_row
        new_max_row = new_real_max_row
        old_max_col = old_real_max_col
        new_max_col = new_real_max_col
        if max_row == 0 or max_col == 0:
            return
        self._buf_log(f" Sheet大小: {max_row}行 x {max_col}列")

        check_value = opts.get('value', True)
        check_formula = opts.get('formula', True)
        check_font = opts.get('font', True)
        check_fill = opts.get('fill', True)
        check_border = opts.get('border', True)
        check_align = opts.get('alignment', True)
        check_nf = opts.get('number_format', True)

        batch_size = 200
        row_count = 0
        for row_idx in range(1, max_row + 1):
            if self.stop_event.is_set():
                raise KeyboardInterrupt
            old_row_data = None
            new_row_data = None
            if row_idx <= old_max_row:
                old_rows = list(old_ws.iter_rows(min_row=row_idx, max_row=row_idx, min_col=1, max_col=max_col, values_only=False))
                old_row_data = old_rows[0] if old_rows else []
            if row_idx <= new_max_row:
                new_rows = list(new_ws.iter_rows(min_row=row_idx, max_row=row_idx, min_col=1, max_col=max_col, values_only=False))
                new_row_data = new_rows[0] if new_rows else []
            for col_idx in range(1, max_col + 1):
                old_cell = old_row_data[col_idx - 1] if old_row_data and col_idx - 1 < len(old_row_data) else None
                new_cell = new_row_data[col_idx - 1] if new_row_data and col_idx - 1 < len(new_row_data) else None
                if old_cell is None and new_cell is None:
                    continue
                if old_cell is None or new_cell is None:
                    old_val = old_cell.value if old_cell else None
                    new_val = new_cell.value if new_cell else None
                    if old_val is None and new_val is None:
                        continue
                    addr = cell_address(col_idx, row_idx)
                    if old_cell is None:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '单元格新增', 'desc': f'新增: {new_val}'})
                    else:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '单元格删除', 'desc': f'删除: {old_val}'})
                    self.stats['diff_cells'] += 1
                    continue
                old_v = old_cell.value
                new_v = new_cell.value
                addr = cell_address(col_idx, row_idx)
                if old_v is None and new_v is None:
                    if not (check_font or check_fill or check_border or check_align or check_nf):
                        continue
                if check_value:
                    val_diff = self._get_cell_diff(old_cell, new_cell)
                    if val_diff:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '内容变化', 'desc': val_diff})
                        self.stats['diff_cells'] += 1
                        continue
                if check_formula and not check_value:
                    old_formula = old_cell.value if isinstance(old_cell.value, str) and old_cell.value.startswith('=') else None
                    new_formula = new_cell.value if isinstance(new_cell.value, str) and new_cell.value.startswith('=') else None
                    if old_formula != new_formula:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '公式变化', 'desc': f"公式: {old_formula} → {new_formula}"})
                        self.stats['diff_cells'] += 1
                        continue
                if check_font:
                    fdiff = self._cmp_font(old_cell.font, new_cell.font)
                    if fdiff:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '字体变化', 'desc': fdiff})
                if check_fill:
                    ffdiff = self._cmp_fill(old_cell.fill, new_cell.fill)
                    if ffdiff:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '填充变化', 'desc': ffdiff})
                if check_border:
                    bdiff = self._cmp_border(old_cell.border, new_cell.border)
                    if bdiff:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '边框变化', 'desc': bdiff})
                if check_align:
                    adiff = self._cmp_alignment(old_cell.alignment, new_cell.alignment)
                    if adiff:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '对齐变化', 'desc': adiff})
                if check_nf:
                    nf1 = old_cell.number_format if old_cell.number_format is not None else 'General'
                    nf2 = new_cell.number_format if new_cell.number_format is not None else 'General'
                    if nf1.strip().lower() != nf2.strip().lower():
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '数字格式变化', 'desc': f'{nf1} → {nf2}'})
                row_count += 1
                if row_count % batch_size == 0:
                    cur_pct = 25 + int(55 * (row_idx / max_row))
                    self.progress(cur_pct, f" {sheet_name}: {row_idx}/{max_row}行...")
                    self._flush_log(force=True)
        if opts.get('rich_text', True):
            old_sheet_rich = self.old_rich.get(sheet_name, {})
            new_sheet_rich = self.new_rich.get(sheet_name, {})
            for ref in set(old_sheet_rich.keys()) | set(new_sheet_rich.keys()):
                rt_diff = compare_rich_text_runs(old_sheet_rich.get(ref), new_sheet_rich.get(ref))
                if rt_diff:
                    self.diffs.append({'sheet': sheet_name, 'address': ref, 'type': '富文本变化', 'desc': rt_diff})
        if opts.get('merged_cells', True):
            self._compare_merged_cells(old_ws, new_ws, sheet_name)
        if opts.get('row_height', True) or opts.get('col_width', True):
            self._compare_row_col_dimensions(old_ws, new_ws, sheet_name)
        if opts.get('images', True):
            self._compare_images(old_ws, new_ws, sheet_name)
        if opts.get('conditional_format', True):
            self._compare_conditional_formats(old_ws, new_ws, sheet_name)

    def _get_cell_diff(self, c1, c2):
        v1 = c1.value
        v2 = c2.value
        opts = self.check_options
        if opts.get('formula', True):
            f1 = isinstance(v1, str) and v1.startswith('=')
            f2 = isinstance(v2, str) and v2.startswith('=')
            if f1 and f2 and v1 == v2:
                return None
        n1 = v1 if v1 is not None else ''
        n2 = v2 if v2 is not None else ''
        f1 = isinstance(v1, str) and v1.startswith('=')
        f2 = isinstance(v2, str) and v2.startswith('=')
        if f1 != f2:
            return f"公式状态: {'是' if f1 else '否'} → {'是' if f2 else '否'}"
        if f1 and f2 and v1 != v2:
            return f"公式: {v1} → {v2}"
        if n1 != n2:
            return f"{str(v1)[:120] if v1 is not None else ''} → {str(v2)[:120] if v2 is not None else ''}"
        return None

    def _cmp_font(self, f1, f2):
        changes = []
        # 仅检查显式字体名称（非None且不为空）
        n1 = f1.name if f1.name else None
        n2 = f2.name if f2.name else None
        # 如果两者都未显式设置，忽略
        if n1 is None and n2 is None:
            pass
        elif n1 is None or n2 is None:
            # 一方显式，另一方默认，忽略（默认渲染相同，不算显式差异）
            pass
        elif n1 != n2:
            changes.append(f"字体: {n1}→{n2}")
        # 仅当显式设置时，检查其他属性
        if n1 is not None and n2 is not None:
            s1 = f1.size if f1.size is not None else 11
            s2 = f2.size if f2.size is not None else 11
            if s1 != s2:
                changes.append(f"字号: {s1}→{s2}")
            b1 = f1.bold if f1.bold is not None else False
            b2 = f2.bold if f2.bold is not None else False
            if b1 != b2:
                changes.append(f"加粗: {b1}→{b2}")
            i1 = f1.italic if f1.italic is not None else False
            i2 = f2.italic if f2.italic is not None else False
            if i1 != i2:
                changes.append(f"斜体: {i1}→{i2}")
            u1 = f1.underline if f1.underline is not None else False
            u2 = f2.underline if f2.underline is not None else False
            if u1 != u2:
                changes.append(f"下划线: {u1}→{u2}")
            c1 = normalize_color_for_compare(f1.color)
            c2 = normalize_color_for_compare(f2.color)
            if c1 != c2:
                if not (isinstance(c1, tuple) and isinstance(c2, tuple) and c1[0] == 'theme' and c1 == c2):
                    changes.append(f"颜色: {rgb_to_hex(f1.color)}→{rgb_to_hex(f2.color)}")
        return '; '.join(changes) if changes else None

    def _cmp_fill(self, f1, f2):
        t1 = f1.fill_type if f1.fill_type is not None else 'none'
        t2 = f2.fill_type if f2.fill_type is not None else 'none'
        sc1 = normalize_color_for_compare(f1.start_color)
        sc2 = normalize_color_for_compare(f2.start_color)
        ec1 = normalize_color_for_compare(f1.end_color)
        ec2 = normalize_color_for_compare(f2.end_color)
        if t1 != t2 or sc1 != sc2 or ec1 != ec2:
            return f"类型: {f1.fill_type}→{f2.fill_type}, 颜色: {rgb_to_hex(f1.start_color)}→{rgb_to_hex(f2.start_color)}"
        return None

    def _cmp_border(self, b1, b2):
        parts = []
        side_names = {'left': '左', 'right': '右', 'top': '上', 'bottom': '下'}
        style_names = {
            'thin': '细线', 'medium': '中等线', 'dashed': '虚线', 'dotted': '点线',
            'double': '双线', 'thick': '粗线', 'dashDot': '点划线', 'dashDotDot': '双点划线',
            'slantDashDot': '斜点划线', 'mediumDashed': '中等虚线', 'mediumDashDot': '中等点划线',
            'mediumDashDotDot': '中等双点划线', 'hair': '发丝线'
        }
        color_names = {
            'FF000000': '黑色', 'FFFFFFFF': '白色', 'FFFF0000': '红色', 'FF00FF00': '绿色',
            'FF0000FF': '蓝色', 'FFFFFF00': '黄色', 'FFFF00FF': '品红', 'FF00FFFF': '青色'
        }
        for side in ['left', 'right', 'top', 'bottom']:
            s1 = getattr(b1, side)
            s2 = getattr(b2, side)
            st1 = s1.style if s1.style is not None else None
            st2 = s2.style if s2.style is not None else None
            c1 = normalize_color_for_compare(s1.color)
            c2 = normalize_color_for_compare(s2.color)
            if st1 != st2 or c1 != c2:
                def color_desc(ci):
                    if ci is None:
                        return '无'
                    if isinstance(ci, tuple):
                        if ci[0] == 'rgb':
                            return color_names.get(ci[1], f"#{ci[1]}")
                        return str(ci)
                    return str(ci)
                def style_desc(st):
                    return style_names.get(st, st or '无')
                parts.append(f"{side_names[side]}: {style_desc(st1)}/{color_desc(c1)}→{style_desc(st2)}/{color_desc(c2)}")
        return '; '.join(parts) if parts else None

    def _cmp_alignment(self, a1, a2):
        changes = []
        h1 = a1.horizontal if a1.horizontal is not None else ''
        h2 = a2.horizontal if a2.horizontal is not None else ''
        if h1 != h2:
            changes.append(f"水平: {h1 or '默认'}→{h2 or '默认'}")
        v1 = a1.vertical if a1.vertical is not None else ''
        v2 = a2.vertical if a2.vertical is not None else ''
        if v1 != v2:
            changes.append(f"垂直: {v1 or '默认'}→{v2 or '默认'}")
        w1 = a1.wrap_text if a1.wrap_text is not None else False
        w2 = a2.wrap_text if a2.wrap_text is not None else False
        if w1 != w2:
            changes.append(f"自动换行: {w1}→{w2}")
        return '; '.join(changes) if changes else None

    def _compare_row_col_dimensions(self, old_ws, new_ws, sheet_name):
        opts = self.check_options
        old_default_row = old_ws.sheet_format.defaultRowHeight
        new_default_row = new_ws.sheet_format.defaultRowHeight
        old_default_col = old_ws.sheet_format.defaultColWidth
        new_default_col = new_ws.sheet_format.defaultColWidth
        if opts.get('row_height', True):
            all_rows = set(old_ws.row_dimensions.keys()) | set(new_ws.row_dimensions.keys())
            for row_idx in all_rows:
                oh = old_ws.row_dimensions[row_idx].height if row_idx in old_ws.row_dimensions else None
                nh = new_ws.row_dimensions[row_idx].height if row_idx in new_ws.row_dimensions else None
                # 处理默认值
                oh_eff = oh if oh is not None else old_default_row
                nh_eff = nh if nh is not None else new_default_row
                if oh_eff is not None and nh_eff is not None and abs(oh_eff - nh_eff) > 0.01:
                    self.diffs.append({'sheet': sheet_name, 'address': f"A{row_idx}", 'type': '行高变化', 'desc': f'行高: {oh} → {nh}'})
        if opts.get('col_width', True):
            all_cols = set(old_ws.column_dimensions.keys()) | set(new_ws.column_dimensions.keys())
            for col_letter in all_cols:
                ow = old_ws.column_dimensions[col_letter].width if col_letter in old_ws.column_dimensions else None
                nw = new_ws.column_dimensions[col_letter].width if col_letter in new_ws.column_dimensions else None
                ow_eff = ow if ow is not None else old_default_col
                nw_eff = nw if nw is not None else new_default_col
                if ow_eff is not None and nw_eff is not None and abs(ow_eff - nw_eff) > 0.01:
                    self.diffs.append({'sheet': sheet_name, 'address': cell_address(column_index_from_string(col_letter), 1), 'type': '列宽变化', 'desc': f'列宽({col_letter}): {ow} → {nw}'})

    def _compare_merged_cells(self, old_ws, new_ws, sheet_name):
        old_merged = set(str(m) for m in old_ws.merged_cells.ranges)
        new_merged = set(str(m) for m in new_ws.merged_cells.ranges)
        for addr in new_merged - old_merged:
            self.diffs.append({'sheet': sheet_name, 'address': addr.split(':')[0], 'type': '合并新增', 'desc': f'新增合并区域 {addr}'})
        for addr in old_merged - new_merged:
            self.diffs.append({'sheet': sheet_name, 'address': addr.split(':')[0], 'type': '合并删除', 'desc': f'删除合并区域 {addr}'})

    def _get_images_from_ws(self, ws):
        images = []
        if hasattr(ws, '_images') and ws._images:
            for img in ws._images:
                try:
                    anchor = img.anchor
                    if anchor and hasattr(anchor, '_from'):
                        col = anchor._from.col + 1
                        row = anchor._from.row + 1
                        images.append((cell_address(col, row), img.width, img.height))
                except:
                    pass
            return sorted(images, key=lambda x: x[0])
        if hasattr(ws, '_drawing') and ws._drawing:
            for anchor in ws._drawing.anchors:
                if hasattr(anchor, 'image'):
                    img = anchor.image
                    col = (anchor._from.col if hasattr(anchor, '_from') else 0) + 1
                    row = (anchor._from.row if hasattr(anchor, '_from') else 0) + 1
                    images.append((cell_address(col, row), img.width, img.height))
            return sorted(images, key=lambda x: x[0])
        return images

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
            if abs(w1 - w2) / max(w1, w2) > 0.01 or abs(h1 - h2) / max(h1, h2) > 0.01:
                if abs(w1 - w2) > 1 or abs(h1 - h2) > 1:
                    changed.append((addr, w1, h1, w2, h2))
        if len(added) + len(removed) + len(changed) > 0:
            for addr in sorted(added):
                self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '图片新增', 'desc': f'新增图片 ({new_set[addr][0]:.0f}x{new_set[addr][1]:.0f})'})
            for addr in sorted(removed):
                self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '图片删除', 'desc': f'删除图片 ({old_set[addr][0]:.0f}x{old_set[addr][1]:.0f})'})
            for addr, w1, h1, w2, h2 in sorted(changed, key=lambda x: x[0]):
                self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '图片尺寸变化', 'desc': f'图片尺寸: {w1:.0f}x{h1:.0f} → {w2:.0f}x{h2:.0f}'})

    def _compare_conditional_formats(self, old_ws, new_ws, sheet_name):
        old_cfs = list(old_ws.conditional_formatting)
        new_cfs = list(new_ws.conditional_formatting)
        old_map = {str(cf.sqref): cf for cf in old_cfs}
        new_map = {str(cf.sqref): cf for cf in new_cfs}
        for rng in set(old_map.keys()) | set(new_map.keys()):
            old_cf = old_map.get(rng)
            new_cf = new_map.get(rng)
            if old_cf is None:
                start = rng.split(':')[0]
                self.diffs.append({'sheet': sheet_name, 'address': start, 'type': '条件格式新增', 'desc': f'新增条件格式范围: {rng}'})
            elif new_cf is None:
                start = rng.split(':')[0]
                self.diffs.append({'sheet': sheet_name, 'address': start, 'type': '条件格式删除', 'desc': f'删除条件格式范围: {rng}'})
            else:
                old_rules = [(r.type, r.priority, str(r.formula), str(r.dxf)) for r in old_cf.rules]
                new_rules = [(r.type, r.priority, str(r.formula), str(r.dxf)) for r in new_cf.rules]
                if old_rules != new_rules:
                    start = rng.split(':')[0]
                    self.diffs.append({'sheet': sheet_name, 'address': start, 'type': '条件格式修改', 'desc': f'条件格式规则变化，范围: {rng}'})

    def _compare_by_check_type(self, check_type, old_cell, new_cell, options=None,
                               old_ws=None, new_ws=None, address=None, sheet_name=''):
        if check_type == 'value':
            if old_cell.value != new_cell.value:
                return f"{old_cell.value} → {new_cell.value}"
        elif check_type == 'formula':
            of = old_cell.value if isinstance(old_cell.value, str) and old_cell.value.startswith('=') else None
            nf = new_cell.value if isinstance(new_cell.value, str) and new_cell.value.startswith('=') else None
            if of != nf:
                return f"公式: {of} → {nf}"
        elif check_type == 'rich_text':
            if old_ws and new_ws:
                old_rich = self.old_rich.get(sheet_name, {}).get(address)
                new_rich = self.new_rich.get(sheet_name, {}).get(address)
                return compare_rich_text_runs(old_rich, new_rich)
        elif check_type == 'font':
            return self._cmp_font(old_cell.font, new_cell.font)
        elif check_type == 'fill':
            return self._cmp_fill(old_cell.fill, new_cell.fill)
        elif check_type == 'border':
            return self._cmp_border(old_cell.border, new_cell.border)
        elif check_type == 'alignment':
            return self._cmp_alignment(old_cell.alignment, new_cell.alignment)
        elif check_type == 'number_format':
            nf1 = old_cell.number_format if old_cell.number_format is not None else 'General'
            nf2 = new_cell.number_format if new_cell.number_format is not None else 'General'
            if nf1.strip().lower() != nf2.strip().lower():
                return f"{nf1} → {nf2}"
        elif check_type == 'merged_cells':
            old_merged = str(old_cell.coordinate) if any(str(old_cell.coordinate) in str(m) for m in old_ws.merged_cells.ranges) else None
            new_merged = str(new_cell.coordinate) if any(str(new_cell.coordinate) in str(m) for m in new_ws.merged_cells.ranges) else None
            if old_merged != new_merged:
                return f"合并区域: {old_merged} → {new_merged}"
        elif check_type == 'row_height':
            oh = old_ws.row_dimensions[old_cell.row].height if old_cell.row in old_ws.row_dimensions else None
            nh = new_ws.row_dimensions[new_cell.row].height if new_cell.row in new_ws.row_dimensions else None
            if oh != nh:
                return f"行高: {oh} → {nh}"
        elif check_type == 'col_width':
            cl = get_column_letter(old_cell.column)
            ow = old_ws.column_dimensions[cl].width if cl in old_ws.column_dimensions else None
            nw = new_ws.column_dimensions[cl].width if cl in new_ws.column_dimensions else None
            if ow != nw:
                return f"列宽({cl}): {ow} → {nw}"
        elif check_type == 'images':
            old_imgs = self._get_images_from_ws(old_ws)
            new_imgs = self._get_images_from_ws(new_ws)
            old_has = any(addr == address for addr, _, _ in old_imgs)
            new_has = any(addr == address for addr, _, _ in new_imgs)
            if old_has != new_has:
                return f"图片存在: {old_has} → {new_has}"
        elif check_type == 'conditional_format':
            old_cf = self._get_conditional_format_for_cell(old_ws, address)
            new_cf = self._get_conditional_format_for_cell(new_ws, address)
            if old_cf != new_cf:
                return f"条件格式: {old_cf} → {new_cf}"
        return None

    def _get_conditional_format_for_cell(self, ws, address):
        for cf in ws.conditional_formatting:
            if address in str(cf.sqref):
                return str(cf.rules)
        return None

    def _compare_sheets(self, old_wb, new_wb):
        old_names = set(old_wb.sheetnames)
        new_names = set(new_wb.sheetnames)
        for name in sorted(new_names - old_names, key=lambda n: list(new_names).index(n)):
            self.sheet_diffs.append({'name': name, 'type': '新增', 'desc': f'新版新增 Sheet: {name}'})
        for name in sorted(old_names - new_names, key=lambda n: list(old_names).index(n)):
            self.sheet_diffs.append({'name': name, 'type': '删除', 'desc': f'旧版有但新版无: {name}'})

    def _row_has_data(self, ws, row, cols):
        for c in range(1, cols + 1):
            if ws.cell(row, c).value is not None:
                return True
        return False

    def _col_has_data(self, ws, col, rows):
        for r in range(1, rows + 1):
            if ws.cell(r, col).value is not None:
                return True
        return False

    def _real_data_range(self, ws):
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0
        if max_row == 0 or max_col == 0:
            return 0, 0, 1, 1
        if max_row <= 50000 and max_col <= 200:
            return 1, 1, max_row, max_col
        check_cols = min(max_col, 500)
        lo, hi = 1, max_row
        real_max_row = 1
        if max_row > 100000 and not self._row_has_data(ws, 100000, check_cols):
            hi = 100000
        if hi > 20000 and not self._row_has_data(ws, 20000, check_cols):
            hi = 20000
        if hi > 5000 and not self._row_has_data(ws, 5000, check_cols):
            hi = 5000
        while lo <= hi:
            mid = (lo + hi) // 2
            if self._row_has_data(ws, mid, check_cols):
                real_max_row = mid
                lo = mid + 1
            else:
                hi = mid - 1
        check_rows = min(real_max_row, 500)
        col_upper = min(max_col, 500)
        if col_upper > 200 and not self._col_has_data(ws, 200, check_rows):
            col_upper = 200
        if col_upper > 50 and not self._col_has_data(ws, 50, check_rows):
            col_upper = 50
        real_max_col = 1
        lo, hi = 1, col_upper
        while lo <= hi:
            mid = (lo + hi) // 2
            if self._col_has_data(ws, mid, check_rows):
                real_max_col = mid
                lo = mid + 1
            else:
                hi = mid - 1
        return 1, 1, real_max_row, real_max_col

# ======================== 对话框 ========================
class CheckOptionsDialog(tb.Toplevel):
    def __init__(self, parent, current_options):
        super().__init__(parent)
        self.title("检测项目设置")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.vars = {}
        main_frame = tb.Frame(self, padding=15)
        main_frame.pack(fill='both', expand=True)
        for col, (group_name, keys) in enumerate(CHECK_OPTION_GROUPS):
            lf = tb.Labelframe(main_frame, text=group_name, padding=(8, 5))
            lf.grid(row=0, column=col, sticky='nsew', padx=5, pady=5)
            for key in keys:
                var = tk.BooleanVar(value=current_options.get(key, True))
                self.vars[key] = var
                cb = tb.Checkbutton(lf, text=CHECK_OPTION_LABELS[key], variable=var, bootstyle="round-toggle")
                cb.pack(anchor='w', pady=2)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=1)
        btn_frame = tb.Frame(main_frame)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=(12, 0))
        tb.Button(btn_frame, text="全选", width=8, command=self._select_all).pack(side='left', padx=(0, 5))
        tb.Button(btn_frame, text="全不选", width=8, command=self._deselect_all).pack(side='left', padx=(0, 5))
        tb.Button(btn_frame, text="取消", width=8, command=self._on_cancel).pack(side='right', padx=(5, 0))
        tb.Button(btn_frame, text="确定", bootstyle=PRIMARY, width=8, command=self._on_ok).pack(side='right')
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        self.geometry(f"+{pw - w // 2}+{ph - h // 2}")
        self.wait_window()

    def _select_all(self):
        for v in self.vars.values():
            v.set(True)

    def _deselect_all(self):
        for v in self.vars.values():
            v.set(False)

    def _on_ok(self):
        self.result = {k: v.get() for k, v in self.vars.items()}
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()

class CheckProjectDialog(tb.Toplevel):
    def __init__(self, parent, old_path, new_path, check_project=None):
        super().__init__(parent)
        self.title("检查项目集配置")
        self.geometry("1100x700")
        self.parent = parent
        self.old_path = old_path
        self.new_path = new_path
        self.project = check_project or CheckProject()
        self.result = None
        self._build_ui()
        self._refresh_rule_list()

    def _build_ui(self):
        info_frame = tb.Labelframe(self, text="项目信息", padding=8)
        info_frame.pack(fill='x', padx=10, pady=5)
        tb.Label(info_frame, text="项目名称:").grid(row=0, column=0, sticky='w', padx=(0, 5))
        self.project_name_var = tk.StringVar(value=self.project.project_name)
        tb.Entry(info_frame, textvariable=self.project_name_var, width=30).grid(row=0, column=1, sticky='w')
        tb.Label(info_frame, text="版本:").grid(row=0, column=2, sticky='w', padx=(15, 5))
        self.version_var = tk.StringVar(value=self.project.version)
        tb.Entry(info_frame, textvariable=self.version_var, width=10).grid(row=0, column=3, sticky='w')
        tb.Label(info_frame, text="描述:").grid(row=0, column=4, sticky='w', padx=(15, 5))
        self.desc_var = tk.StringVar(value=self.project.description)
        tb.Entry(info_frame, textvariable=self.desc_var, width=30).grid(row=0, column=5, sticky='w')

        list_frame = tb.Labelframe(self, text="规则列表", padding=8)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        columns = ('rule_name', 'data_source', 'checks')
        self.tree = tb.Treeview(list_frame, columns=columns, show='headings', height=15)
        self.tree.heading('rule_name', text='规则名称')
        self.tree.heading('data_source', text='数据源')
        self.tree.heading('checks', text='检查项数')
        self.tree.column('rule_name', width=200)
        self.tree.column('data_source', width=300)
        self.tree.column('checks', width=80, anchor='center')
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar = tb.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=scrollbar.set)

        btn_frame = tb.Frame(self, padding=5)
        btn_frame.pack(fill='x', padx=10, pady=5)
        tb.Button(btn_frame, text="添加规则", bootstyle=PRIMARY, command=self.add_rule).pack(side='left', padx=5)
        tb.Button(btn_frame, text="复制规则", bootstyle="secondary", command=self.copy_rule).pack(side='left', padx=5)
        tb.Button(btn_frame, text="编辑规则", bootstyle=INFO, command=self.edit_rule).pack(side='left', padx=5)
        tb.Button(btn_frame, text="删除规则", bootstyle=DANGER, command=self.delete_rule).pack(side='left', padx=5)
        tb.Button(btn_frame, text="上移", bootstyle="outline", command=lambda: self.move_rule(-1)).pack(side='left', padx=5)
        tb.Button(btn_frame, text="下移", bootstyle="outline", command=lambda: self.move_rule(1)).pack(side='left', padx=5)
        tb.Separator(btn_frame, orient='vertical').pack(side='left', fill='y', padx=10)
        tb.Button(btn_frame, text="保存项目", bootstyle=SUCCESS, command=self.save_project).pack(side='left', padx=5)
        tb.Button(btn_frame, text="加载项目", bootstyle="outline", command=self.load_project).pack(side='left', padx=5)
        tb.Button(btn_frame, text="应用", bootstyle=INFO, command=self.apply_project).pack(side='right', padx=5)

    def _refresh_rule_list(self):
        self.tree.delete(*self.tree.get_children())
        for idx, rule in enumerate(self.project.rules):
            ds = rule.data_source
            ds_str = f"{ds.get('sheet', '?')} | {ds.get('anchor', {}).get('text', '?')}"
            checks_count = len(rule.checks)
            self.tree.insert('', 'end', iid=str(idx), values=(rule.rule_name, ds_str, checks_count))

    def add_rule(self):
        if not os.path.isfile(self.old_path) and not os.path.isfile(self.new_path):
            messagebox.showwarning("提示", "请先在主界面选择有效的Excel文件")
            return
        dlg = RuleEditorDialog(self, self.old_path, self.new_path)
        if dlg.result is not None:
            self.project.rules.append(dlg.result)
            self._refresh_rule_list()

    def copy_rule(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择要复制的规则")
            return
        idx = int(sel[0])
        original = self.project.rules[idx]
        new_rule = copy.deepcopy(original)
        new_rule.rule_name = original.rule_name + "_副本"
        self.project.rules.append(new_rule)
        self._refresh_rule_list()
        self.tree.selection_set(str(len(self.project.rules) - 1))

    def edit_rule(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择要编辑的规则")
            return
        idx = int(sel[0])
        rule = self.project.rules[idx]
        dlg = RuleEditorDialog(self, self.old_path, self.new_path, rule=rule)
        if dlg.result is not None:
            self.project.rules[idx] = dlg.result
            self._refresh_rule_list()

    def delete_rule(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择要删除的规则")
            return
        idx = int(sel[0])
        if messagebox.askyesno("确认", "确定删除该规则？"):
            self.project.rules.pop(idx)
            self._refresh_rule_list()

    def move_rule(self, direction):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        new_idx = idx + direction
        if 0 <= new_idx < len(self.project.rules):
            self.project.rules[idx], self.project.rules[new_idx] = self.project.rules[new_idx], self.project.rules[idx]
            self._refresh_rule_list()
            self.tree.selection_set(str(new_idx))

    def save_project(self):
        self.project.project_name = self.project_name_var.get()
        self.project.version = self.version_var.get()
        self.project.description = self.desc_var.get()
        if not self.project.project_name:
            messagebox.showwarning("提示", "项目名称不能为空")
            return
        # 默认保存到程序目录
        filepath = os.path.join(PROGRAM_DIR, f"{self.project.project_name}.json")
        if os.path.exists(filepath):
            if not messagebox.askyesno("提示", "文件已存在，是否覆盖？"):
                return
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(self.project.to_dict(), f, ensure_ascii=False, indent=2)
        messagebox.showinfo("成功", f"项目已保存到: {filepath}")

    def load_project(self):
        # 列出程序目录下所有 .json 文件供选择
        json_files = [f for f in os.listdir(PROGRAM_DIR) if f.endswith('.json') and f != 'main.py']
        if not json_files:
            messagebox.showwarning("提示", "程序目录下没有规则文件")
            return
        # 弹出一个选择对话框（简化：使用filedialog限制目录）
        filepath = filedialog.askopenfilename(
            initialdir=PROGRAM_DIR,
            title="选择规则文件",
            filetypes=[("JSON files", "*.json")]
        )
        if not filepath:
            return
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.project = CheckProject.from_dict(data)
            self.project_name_var.set(self.project.project_name)
            self.version_var.set(self.project.version)
            self.desc_var.set(self.project.description)
            self._refresh_rule_list()
            messagebox.showinfo("成功", "项目已加载")
        except Exception as e:
            messagebox.showerror("错误", f"加载失败：{str(e)}")

    def apply_project(self):
        self.project.project_name = self.project_name_var.get()
        self.project.version = self.version_var.get()
        self.project.description = self.desc_var.get()
        self.result = self.project
        self.destroy()

class RuleEditorDialog(tb.Toplevel):
    def __init__(self, parent, old_path, new_path, rule=None):
        super().__init__(parent)
        self.title("编辑规则")
        self.geometry("1100x800")  # 增大高度
        self.parent = parent
        self.old_path = old_path
        self.new_path = new_path
        self.result = None
        self.rule = rule if rule else CheckRule()
        self._build_ui()
        self._load_rule_data()
        self.wait_window()

    def _get_sheet_names(self):
        # 优先从旧版读取，如果没有再从新版读取
        for p in [self.old_path, self.new_path]:
            if os.path.isfile(p):
                sheets = get_sheet_names_fast(p)
                if sheets:
                    return sheets
        return []

    def _build_ui(self):
        main_frame = tb.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)
        left_frame = tb.Frame(main_frame, width=420)
        left_frame.pack(side='left', fill='y', padx=(0, 10))
        left_frame.pack_propagate(False)
        ds_frame = tb.Labelframe(left_frame, text="数据源配置", padding=10)
        ds_frame.pack(fill='both', expand=True, pady=5)

        tb.Label(ds_frame, text="规则名称:").pack(anchor='w')
        self.rule_name_var = tk.StringVar()
        tb.Entry(ds_frame, textvariable=self.rule_name_var, width=40).pack(fill='x', pady=2)

        tb.Label(ds_frame, text="Sheet:").pack(anchor='w')
        sheet_row = tb.Frame(ds_frame)
        sheet_row.pack(fill='x', pady=2)
        self.sheet_var = tk.StringVar()
        self.sheet_cb = tb.Combobox(sheet_row, textvariable=self.sheet_var, width=30)
        self.sheet_cb.pack(side='left', fill='x', expand=True)
        tb.Button(sheet_row, text="抓取", bootstyle="outline", width=5, command=self.fetch_sheets).pack(side='right', padx=(3, 0))

        tb.Label(ds_frame, text="锚点文字:").pack(anchor='w')
        self.anchor_text_var = tk.StringVar()
        tb.Entry(ds_frame, textvariable=self.anchor_text_var, width=40).pack(fill='x', pady=2)

        tb.Label(ds_frame, text="搜索范围:").pack(anchor='w')
        self.search_in_var = tk.StringVar(value='all')
        tb.Combobox(ds_frame, textvariable=self.search_in_var, values=['all', 'first_row', 'first_col'], width=38).pack(fill='x', pady=2)

        tb.Label(ds_frame, text="模式:").pack(anchor='w')
        self.mode_var = tk.StringVar(value='offset')
        self.mode_cb = tb.Combobox(ds_frame, textvariable=self.mode_var, values=['offset', 'collect', 'intersection', 'range'], width=38)
        self.mode_cb.pack(fill='x', pady=2)
        self.mode_cb.bind('<<ComboboxSelected>>', lambda e: self._build_param_fields())

        self.param_frame = tb.Frame(ds_frame)
        self.param_frame.pack(fill='x', pady=2)
        self._build_param_fields()

        # 底部按钮放在left_frame底部，避免被压缩
        btn_frame = tb.Frame(left_frame)
        btn_frame.pack(side='bottom', fill='x', pady=5)
        tb.Button(btn_frame, text="确定", bootstyle=PRIMARY, width=8, command=self.on_ok).pack(side='left', padx=5)
        tb.Button(btn_frame, text="取消", bootstyle="outline", width=8, command=self.on_cancel).pack(side='right', padx=5)

        right_frame = tb.Frame(main_frame)
        right_frame.pack(side='right', fill='both', expand=True)
        check_frame = tb.Labelframe(right_frame, text="检查项（可多选）", padding=10)
        check_frame.pack(fill='both', expand=True)
        canvas = tk.Canvas(check_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(check_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tb.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.check_vars = {}
        self.expect_vars = {}
        for group_name, keys in CHECK_OPTION_GROUPS:
            lf = tb.Labelframe(scrollable_frame, text=group_name, padding=(8, 5))
            lf.pack(fill='x', pady=(0, 8), padx=5)
            for key in keys:
                row = tb.Frame(lf)
                row.pack(fill='x', pady=1)
                left_cell = tb.Frame(row)
                left_cell.pack(side='left', fill='x', expand=True)
                var = tk.BooleanVar(value=True)
                self.check_vars[key] = var
                cb = tb.Checkbutton(left_cell, text=CHECK_OPTION_LABELS[key], variable=var, bootstyle="round-toggle")
                cb.pack(side='left', anchor='w')
                right_cell = tb.Frame(row)
                right_cell.pack(side='right')
                tb.Label(right_cell, text="期望:").pack(side='left', padx=(10, 2))
                expect_var = tk.StringVar(value='same')
                tb.Combobox(right_cell, textvariable=expect_var, values=['same', 'different'], width=8).pack(side='left')
                self.expect_vars[key] = expect_var

    def fetch_sheets(self):
        sheets = self._get_sheet_names()
        if sheets:
            self.sheet_cb['values'] = sheets
            if not self.sheet_var.get() and sheets:
                self.sheet_var.set(sheets[0])
        else:
            messagebox.showwarning("提示", "无法读取Sheet列表，请检查文件路径")

    def _build_param_fields(self):
        for widget in self.param_frame.winfo_children():
            widget.destroy()
        mode = self.mode_var.get()
        if mode == 'offset':
            tb.Label(self.param_frame, text="行偏移:").pack(anchor='w')
            self.offset_row_var = tk.StringVar(value='0')
            tb.Entry(self.param_frame, textvariable=self.offset_row_var, width=10).pack(fill='x', pady=2)
            tb.Label(self.param_frame, text="列偏移:").pack(anchor='w')
            self.offset_col_var = tk.StringVar(value='0')
            tb.Entry(self.param_frame, textvariable=self.offset_col_var, width=10).pack(fill='x', pady=2)
        elif mode == 'collect':
            tb.Label(self.param_frame, text="方向:").pack(anchor='w')
            self.collect_dir_var = tk.StringVar(value='down')
            tb.Combobox(self.param_frame, textvariable=self.collect_dir_var, values=['down', 'right'], width=10).pack(fill='x', pady=2)
            tb.Label(self.param_frame, text="起始偏移:").pack(anchor='w')
            self.collect_start_var = tk.StringVar(value='1')
            tb.Entry(self.param_frame, textvariable=self.collect_start_var, width=10).pack(fill='x', pady=2)
            tb.Label(self.param_frame, text="最大数量:").pack(anchor='w')
            self.collect_max_var = tk.StringVar(value='100')
            tb.Entry(self.param_frame, textvariable=self.collect_max_var, width=10).pack(fill='x', pady=2)
        elif mode == 'intersection':
            tb.Label(self.param_frame, text="行锚点文字:").pack(anchor='w')
            self.row_anchor_text_var = tk.StringVar()
            tb.Entry(self.param_frame, textvariable=self.row_anchor_text_var, width=20).pack(fill='x', pady=2)
            tb.Label(self.param_frame, text="列锚点文字:").pack(anchor='w')
            self.col_anchor_text_var = tk.StringVar()
            tb.Entry(self.param_frame, textvariable=self.col_anchor_text_var, width=20).pack(fill='x', pady=2)
        elif mode == 'range':
            tb.Label(self.param_frame, text="行偏移:").pack(anchor='w')
            self.range_row_offset_var = tk.StringVar(value='0')
            tb.Entry(self.param_frame, textvariable=self.range_row_offset_var, width=10).pack(fill='x', pady=2)
            tb.Label(self.param_frame, text="列偏移:").pack(anchor='w')
            self.range_col_offset_var = tk.StringVar(value='0')
            tb.Entry(self.param_frame, textvariable=self.range_col_offset_var, width=10).pack(fill='x', pady=2)
            tb.Label(self.param_frame, text="行数:").pack(anchor='w')
            self.range_row_count_var = tk.StringVar(value='1')
            tb.Entry(self.param_frame, textvariable=self.range_row_count_var, width=10).pack(fill='x', pady=2)
            tb.Label(self.param_frame, text="列数:").pack(anchor='w')
            self.range_col_count_var = tk.StringVar(value='1')
            tb.Entry(self.param_frame, textvariable=self.range_col_count_var, width=10).pack(fill='x', pady=2)
            tb.Label(self.param_frame, text="排除单元格(逗号分隔，如A5,C7或[1,0]):").pack(anchor='w')
            self.range_exclude_var = tk.StringVar(value='')
            tb.Entry(self.param_frame, textvariable=self.range_exclude_var, width=30).pack(fill='x', pady=2)

    def _load_rule_data(self):
        self.rule_name_var.set(self.rule.rule_name)
        ds = self.rule.data_source
        self.sheet_var.set(ds.get('sheet', ''))
        self.anchor_text_var.set(ds.get('anchor', {}).get('text', ''))
        self.search_in_var.set(ds.get('search_in', 'all'))
        self.mode_var.set(ds.get('mode', 'offset'))
        self._build_param_fields()
        target = ds.get('target', {})
        mode = self.mode_var.get()
        if mode == 'offset':
            self.offset_row_var.set(str(target.get('row_offset', 0)))
            self.offset_col_var.set(str(target.get('col_offset', 0)))
        elif mode == 'collect':
            self.collect_dir_var.set(target.get('direction', 'down'))
            self.collect_start_var.set(str(target.get('start_offset', 1)))
            self.collect_max_var.set(str(target.get('max_count', 100)))
        elif mode == 'range':
            self.range_row_offset_var.set(str(target.get('row_offset', 0)))
            self.range_col_offset_var.set(str(target.get('col_offset', 0)))
            self.range_row_count_var.set(str(target.get('row_count', 1)))
            self.range_col_count_var.set(str(target.get('col_count', 1)))
            ex_list = target.get('exclude', [])
            self.range_exclude_var.set(','.join(str(x) if isinstance(x, str) else f"[{x[0]},{x[1]}]" for x in ex_list))
        elif mode == 'intersection':
            row_anchor = ds.get('row_anchor', {})
            col_anchor = ds.get('col_anchor', {})
            self.row_anchor_text_var.set(row_anchor.get('text', '').strip())
            self.col_anchor_text_var.set(col_anchor.get('text', '').strip())
        for check in self.rule.checks:
            if check.check_type in self.check_vars:
                self.check_vars[check.check_type].set(check.enabled)
                self.expect_vars[check.check_type].set(check.expect)

    def on_ok(self):
        self.rule.rule_name = self.rule_name_var.get()
        ds = {
            'name': self.rule_name_var.get(),
            'sheet': self.sheet_var.get(),
            'anchor': {'text': self.anchor_text_var.get()},
            'search_in': self.search_in_var.get(),
            'mode': self.mode_var.get()
        }
        if self.mode_var.get() == 'offset':
            ds['target'] = {
                'row_offset': int(self.offset_row_var.get()),
                'col_offset': int(self.offset_col_var.get())
            }
        elif self.mode_var.get() == 'collect':
            ds['target'] = {
                'direction': self.collect_dir_var.get(),
                'start_offset': int(self.collect_start_var.get()),
                'max_count': int(self.collect_max_var.get())
            }
        elif self.mode_var.get() == 'intersection':
            ds['row_anchor'] = {'text': self.row_anchor_text_var.get().strip(), 'search_in': 'all'}
            ds['col_anchor'] = {'text': self.col_anchor_text_var.get().strip(), 'search_in': 'all'}
        elif self.mode_var.get() == 'range':
            ds['target'] = {
                'row_offset': int(self.range_row_offset_var.get()),
                'col_offset': int(self.range_col_offset_var.get()),
                'row_count': int(self.range_row_count_var.get()),
                'col_count': int(self.range_col_count_var.get()),
                'exclude': self._parse_exclude(self.range_exclude_var.get())
            }
        self.rule.data_source = ds
        checks = []
        for key, var in self.check_vars.items():
            if var.get():
                checks.append(CheckItemConfig(check_type=key, enabled=True, expect=self.expect_vars[key].get()))
        self.rule.checks = checks
        self.result = self.rule
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

    def _parse_exclude(self, text):
        if not text.strip():
            return []
        parts = [p.strip() for p in text.split(',') if p.strip()]
        exclude = []
        for p in parts:
            if re.match(r'^[A-Za-z]+\d+$', p):
                exclude.append(p)
            elif re.match(r'^\[\d+,\d+\]$', p):
                inner = p[1:-1].split(',')
                exclude.append([int(inner[0]), int(inner[1])])
        return exclude

# ======================== 主界面 ========================
class DiffViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel报告检查工具")
        self.root.geometry("1100x750")
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.topmost = tk.BooleanVar(value=False)
        self.check_options = dict(DEFAULT_CHECK_OPTIONS)
        self.plugin_manager = None
        self.config_file = None
        self.stop_event = threading.Event()
        self.check_project = None

        toolbar = tb.Frame(root, padding=5)
        toolbar.pack(fill='x')
        toolbar.columnconfigure(6, weight=1)

        self.start_btn = tb.Button(toolbar, text="常规差异对比", bootstyle=INFO, width=12, command=self.start_compare)
        self.start_btn.grid(row=0, column=0, rowspan=2, sticky='nsew', padx=2, pady=1)

        self.stop_btn = tb.Button(toolbar, text="停止检查", bootstyle=DANGER, width=8, command=self.stop_compare, state='disabled')
        self.stop_btn.grid(row=0, column=1, rowspan=2, sticky='nsew', padx=2, pady=1)

        self.settings_btn = tb.Button(toolbar, text="常规差异检测设置", bootstyle="outline", width=14, command=self.open_check_options)
        self.settings_btn.grid(row=0, column=2, rowspan=2, sticky='nsew', padx=2, pady=1)

        self.project_btn = tb.Button(toolbar, text="进阶检查规则", bootstyle="outline-primary", width=12, command=self.open_project_dialog)
        self.project_btn.grid(row=0, column=3, rowspan=2, sticky='nsew', padx=2, pady=1)

        self.config_btn = tb.Button(toolbar, text="导入规则", bootstyle="outline-primary", width=10, command=self.load_check_project)
        self.config_btn.grid(row=0, column=4, rowspan=2, sticky='nsew', padx=2, pady=1)

        tb.Separator(toolbar, orient='vertical').grid(row=0, column=5, rowspan=2, sticky='ns', padx=8)

        path_frame = tb.Frame(toolbar)
        path_frame.grid(row=0, column=6, rowspan=2, sticky='nsew', padx=(0, 10))
        path_frame.columnconfigure(1, weight=1)

        tb.Label(path_frame, text="旧版:").grid(row=0, column=0, sticky='w', padx=(0, 5), pady=2)
        self.old_entry = tb.Entry(path_frame, textvariable=self.old_path)
        self.old_entry.grid(row=0, column=1, sticky='ew', pady=2)
        tb.Button(path_frame, text="浏览", bootstyle=PRIMARY, width=6, command=lambda: self.browse(self.old_path)).grid(row=0, column=2, padx=(5, 0), pady=2)

        tb.Label(path_frame, text="新版:").grid(row=1, column=0, sticky='w', padx=(0, 5), pady=2)
        self.new_entry = tb.Entry(path_frame, textvariable=self.new_path)
        self.new_entry.grid(row=1, column=1, sticky='ew', pady=2)
        tb.Button(path_frame, text="浏览", bootstyle=PRIMARY, width=6, command=lambda: self.browse(self.new_path)).grid(row=1, column=2, padx=(5, 0), pady=2)

        tb.Checkbutton(toolbar, text="置顶", variable=self.topmost, command=self.toggle_topmost, bootstyle="round-toggle").grid(
            row=0, column=7, rowspan=2, padx=(0, 5), pady=2, sticky='w')

        self.progress = tb.Progressbar(root, mode='determinate', bootstyle=PRIMARY)
        self.progress.pack(fill='x', padx=5, pady=(0, 5))

        # Treeview 区域，使用 grid 布局保证滚动条始终显示
        tree_frame = tb.Frame(root, padding=(5, 0))
        tree_frame.pack(fill='both', expand=True)
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.tree = tb.Treeview(tree_frame, columns=('action', 'address', 'type'), show='tree headings', bootstyle=PRIMARY)
        self.tree.heading('#0', text='Sheet / 差异项')
        self.tree.heading('action', text='')
        self.tree.heading('address', text='位置')
        self.tree.heading('type', text='类型')
        self.tree.column('#0', width=250)
        self.tree.column('action', width=40, anchor='center', stretch=False)
        self.tree.column('address', width=80)
        self.tree.column('type', width=100)
        # 配置标签颜色
        self.tree.tag_configure('sheet', foreground='blue', font=('微软雅黑', 10, 'bold'))

        scroll_y = tb.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview, bootstyle=ROUND)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')

        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Double-1>', self.on_tree_double_click)
        # 绑定点击action列事件
        self.tree.bind('<Button-1>', self.on_tree_click)

        bottom_frame = tb.Frame(root, padding=5)
        bottom_frame.pack(fill='x')
        bottom_frame.columnconfigure(0, weight=0)
        bottom_frame.columnconfigure(1, weight=1)

        detailf = tb.Labelframe(bottom_frame, text="差异详情", padding=5, bootstyle=INFO)
        detailf.grid(row=0, column=0, sticky='nsew', padx=(0, 3))
        self.detail = tk.Text(detailf, width=42, height=6, wrap='word', font=("微软雅黑", 9),
                              bg='#ffffff', fg='#212529', relief='flat',
                              highlightthickness=1, highlightbackground='#dee2e6',
                              highlightcolor='#0d6efd', padx=5, pady=5)
        self.detail.pack(fill='both', expand=True)

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
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p:
            var.set(p)

    def toggle_topmost(self):
        self.root.attributes('-topmost', self.topmost.get())

    def open_check_options(self):
        if self.check_project:
            messagebox.showinfo("提示", "当前为规则检查模式，无需设置常规差异检测")
            return
        dlg = CheckOptionsDialog(self.root, self.check_options)
        if dlg.result is not None:
            self.check_options = dlg.result
            enabled_count = sum(1 for v in self.check_options.values() if v)
            total_count = len(self.check_options)
            self.settings_btn.configure(text=f"常规差异检测设置 ({enabled_count}/{total_count})")
            self.log(f"检测设置已更新：{enabled_count}/{total_count} 项已启用")

    def open_project_dialog(self):
        old = self.old_path.get()
        new = self.new_path.get()
        if not old:
            messagebox.showwarning("提示", "请先选择旧版文件路径")
            return
        dlg = CheckProjectDialog(self.root, old, new, self.check_project)
        if dlg.result is not None:
            self.check_project = dlg.result
            self.start_btn.configure(text=f"规则检查\n（{self.check_project.project_name}）", width=16)
            self.settings_btn.configure(state='disabled')
            self.project_btn.configure(text="进阶检查规则")
            self.log(f"检查项目已加载：{self.check_project.project_name}，包含 {len(self.check_project.rules)} 条规则")

    def clear_check_project(self):
        self.check_project = None
        self.start_btn.configure(text="常规差异对比", width=12)
        self.settings_btn.configure(state='normal')
        self.config_btn.configure(text="导入规则", bootstyle="outline-primary", command=self.load_check_project)
        self.project_btn.configure(text="进阶检查规则", command=self.open_project_dialog)
        self.tree.delete(*self.tree.get_children())
        self.detail.delete('1.0', 'end')
        self.diff_items = []
        self.result_data = None
        self.log("已退出规则检查模式，恢复常规差异对比")

    def load_check_project(self):
        # 列出程序目录下所有 .json 文件供选择
        json_files = [f for f in os.listdir(PROGRAM_DIR) if f.endswith('.json')]
        if not json_files:
            messagebox.showwarning("提示", "程序目录下没有规则文件")
            return
        filepath = filedialog.askopenfilename(
            initialdir=PROGRAM_DIR,
            title="选择规则文件",
            filetypes=[("JSON files", "*.json")]
        )
        if not filepath:
            return
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.check_project = CheckProject.from_dict(data)
            self.start_btn.configure(text=f"规则检查\n（{self.check_project.project_name}）", width=16)
            self.settings_btn.configure(state='disabled')
            self.config_btn.configure(text="退出规则模式", bootstyle=DANGER, command=self.clear_check_project)
            self.log(f"检查项目已加载：{self.check_project.project_name}，包含 {len(self.check_project.rules)} 条规则")
        except Exception as e:
            messagebox.showerror("错误", f"加载检查项目失败：{str(e)}")

    def log(self, msg):
        def _log():
            self.log_text.insert('end', f"{msg}\n")
            self.log_text.see('end')
            self.root.update_idletasks()
        self.root.after(0, _log)

    def update_progress(self, val, stat=""):
        def _update():
            self.progress['value'] = val
        self.root.after(0, _update)

    def set_progress_mode(self, mode):
        def _switch():
            self.progress.configure(mode=mode)
            if mode == 'indeterminate':
                self.progress.start(50)
            else:
                self.progress.stop()
        self.root.after(0, _switch)

    def start_compare(self):
        old = self.old_path.get()
        new = self.new_path.get()
        if not old or not new:
            messagebox.showerror("错误", "请选择两个文件")
            return
        if not os.path.isfile(old) or not os.path.isfile(new):
            messagebox.showerror("错误", "文件不存在")
            return
        if not old.lower().endswith('.xlsx') or not new.lower().endswith('.xlsx'):
            messagebox.showerror("错误", "仅支持 .xlsx 格式文件")
            return
        self.stop_event.clear()
        self.start_btn.configure(state='disabled')
        self.stop_btn.configure(state='normal')
        self.tree.delete(*self.tree.get_children())
        self.detail.delete('1.0', 'end')
        self.diff_items = []

        # 添加日志分隔符
        self.log("=" * 50)
        self.log("开始新的检查")
        self.log("=" * 50)

        current_opts = dict(self.check_options)
        pm = self.plugin_manager
        cp = self.check_project

        def worker():
            comparer = None
            try:
                comparer = OpenpyxlComparer(old, new, self.log, self.update_progress,
                                            check_options=current_opts, plugin_manager=pm,
                                            progress_mode_fn=self.set_progress_mode,
                                            check_project=cp,
                                            stop_event=self.stop_event,
                                            mode='diff')
                comparer.run()
                self.result_data = (comparer.diffs, comparer.sheet_diffs, comparer.stats)
                self.root.after(0, self.populate_tree)
            except KeyboardInterrupt:
                if comparer:
                    self.result_data = (comparer.diffs, comparer.sheet_diffs, comparer.stats)
                else:
                    self.result_data = ([], [], {'total_cells': 0, 'diff_cells': 0, 'added_sheets': [], 'removed_sheets': [], 'images_diff': 0})
                self.root.after(0, self.populate_tree)
            except Exception as e:
                import traceback
                tb_str = traceback.format_exc()
                error_msg = f"{str(e)}\n\n{tb_str}"
                self.root.after(0, lambda msg=error_msg: messagebox.showerror("对比失败", msg))
            finally:
                self.root.after(0, self.on_comparison_finished)
        threading.Thread(target=worker, daemon=True).start()

    def on_comparison_finished(self):
        self.start_btn.configure(state='normal')
        self.stop_btn.configure(state='disabled', bootstyle=DANGER)
        if self.stop_event.is_set():
            self.log("检查已停止")
        self.progress['value'] = 100

    def populate_tree(self):
        diffs, sheet_diffs, stats = self.result_data
        normal_diffs = []
        rule_pass_diffs = []
        plugin_diffs = []
        for d in diffs:
            if d['sheet'] == '🔍 数据检查':
                plugin_diffs.append(d)
            elif d.get('rule_pass'):
                rule_pass_diffs.append(d)
            else:
                normal_diffs.append(d)

        if sheet_diffs:
            sn = self.tree.insert('', 'end', text='📋 Sheet 结构差异', open=True, tags=('sheet',))
            for sd in sheet_diffs:
                node = self.tree.insert(sn, 'end', text=sd['desc'], values=('', sd['name'], sd['type']))
                self.diff_items.append((node, {'type': 'sheet_struct', 'data': sd}))

        if plugin_diffs:
            pn = self.tree.insert('', 'end', text='🔍 数据检查结果', open=True)
            for d in plugin_diffs:
                node = self.tree.insert(pn, 'end', text=d['desc'][:80], values=('', d['address'], d['type']))
                self.diff_items.append((node, {'type': 'cell', 'data': d}))

        if normal_diffs:
            pn = self.tree.insert('', 'end', text='⚠️ 需人工复核（未豁免）', open=True, tags=('sheet',))
            dmap = {}
            for d in normal_diffs:
                dmap.setdefault(d['sheet'], []).append(d)
            # 按旧版sheet顺序
            for sname in self.old_sheet_order:
                if sname in dmap:
                    sn = self.tree.insert(pn, 'end', text=f"📄 {sname}", open=True, tags=('sheet',))
                    for d in dmap[sname]:
                        node = self.tree.insert(sn, 'end', text=d['desc'][:80], values=('', d['address'], d['type']))
                        self.diff_items.append((node, {'type': 'cell', 'data': d}))
            # 处理新增或不在旧版中的sheet（排在最后）
            for sname in dmap:
                if sname not in self.old_sheet_order:
                    sn = self.tree.insert(pn, 'end', text=f"📄 {sname}", open=True, tags=('sheet',))
                    for d in dmap[sname]:
                        node = self.tree.insert(sn, 'end', text=d['desc'][:80], values=('', d['address'], d['type']))
                        self.diff_items.append((node, {'type': 'cell', 'data': d}))

        if rule_pass_diffs:
            pn = self.tree.insert('', 'end', text='✅ 已豁免（可展开）', open=False, tags=('sheet',))
            dmap = {}
            for d in rule_pass_diffs:
                dmap.setdefault(d['sheet'], []).append(d)
            for sname in self.old_sheet_order:
                if sname in dmap:
                    sn = self.tree.insert(pn, 'end', text=f"📄 {sname}", open=True, tags=('sheet',))
                    for d in dmap[sname]:
                        desc = d['desc']
                        if d.get('rule_name'):
                            desc += f" [规则: {d['rule_name']}]"
                        node = self.tree.insert(sn, 'end', text=desc[:100], values=('', d['address'], d['type']))
                        self.diff_items.append((node, {'type': 'cell', 'data': d}))
            for sname in dmap:
                if sname not in self.old_sheet_order:
                    sn = self.tree.insert(pn, 'end', text=f"📄 {sname}", open=True, tags=('sheet',))
                    for d in dmap[sname]:
                        desc = d['desc']
                        if d.get('rule_name'):
                            desc += f" [规则: {d['rule_name']}]"
                        node = self.tree.insert(sn, 'end', text=desc[:100], values=('', d['address'], d['type']))
                        self.diff_items.append((node, {'type': 'cell', 'data': d}))

        enabled = sum(1 for v in self.check_options.values() if v)
        total = len(self.check_options)
        self.log(f"检查完成：需人工复核 {len(normal_diffs)} 项，已豁免 {len(rule_pass_diffs)} 项，插件 {len(plugin_diffs)} 项")

    def on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        node = sel[0]
        target = next((d for n, d in self.diff_items if n == node), None)
        if not target:
            return
        if target['type'] == 'cell' or target['type'] == 'sheet_struct':
            d = target['data']
            self.detail.delete('1.0', 'end')
            self.detail.insert('1.0', f"Sheet: {d['sheet']}\n位置: {d.get('address', '?')}\n类型: {d['type']}\n描述: {d['desc']}")

    def on_tree_click(self, event):
        # 判断点击的是否为action列
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        col = self.tree.identify_column(event.x)
        if col != '#1':  # action列是第一列（columns中第一个）
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        # 如果是子节点，则收起其父节点；如果是父节点，则收起自身
        parent = self.tree.parent(iid)
        if parent:
            self.tree.item(parent, open=False)
        else:
            self.tree.item(iid, open=False)

    def jump_to_excel(self, file_path, sheet_name, cell_addr):
        try:
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except pythoncom.com_error:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = True
            target_wb = None
            for wb in excel.Workbooks:
                if normalize_path(wb.FullName) == normalize_path(file_path):
                    target_wb = wb
                    break
            if target_wb is None:
                return False, f"文件未在 Excel 中打开：{file_path}"
            target_wb.Activate()
            try:
                ws = target_wb.Worksheets(sheet_name)
            except Exception as e:
                return False, f"工作表不存在：{sheet_name}（{e}）"
            ws.Activate()
            col_str = ''.join(ch for ch in cell_addr if ch.isalpha())
            row_str = ''.join(ch for ch in cell_addr if ch.isdigit())
            if not col_str or not row_str:
                return False, f"无效单元格地址：{cell_addr}"
            col_idx = column_index_from_string(col_str)
            row_idx = int(row_str)
            excel.ActiveWindow.ScrollRow = row_idx
            excel.ActiveWindow.ScrollColumn = col_idx
            ws.Range(cell_addr).Select()
            return True, None
        except Exception as e:
            return False, str(e)

    def on_tree_double_click(self, event):
        # 双击跳转（除非是action列）
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        col = self.tree.identify_column(event.x)
        if col == '#1':  # action列不跳转
            return
        sel = self.tree.selection()
        if not sel:
            return
        node = sel[0]
        item_info = next((d for n, d in self.diff_items if n == node), None)
        if not item_info or item_info['type'] != 'cell':
            return
        data = item_info['data']
        sheet_name = data.get('sheet')
        cell_addr = data.get('address')
        if not sheet_name or not cell_addr:
            return
        if ':' in cell_addr:
            cell_addr = cell_addr.split(':')[0]
        if not re.match(r'^[A-Za-z]+[0-9]+$', cell_addr):
            self.log(f"跳过跳转：无效地址 {cell_addr}")
            return
        old_path = self.old_path.get()
        new_path = self.new_path.get()
        old_ok, old_err = self.jump_to_excel(old_path, sheet_name, cell_addr)
        new_ok, new_err = self.jump_to_excel(new_path, sheet_name, cell_addr)
        if not old_ok or not new_ok:
            msgs = []
            if not old_ok:
                msgs.append(f"旧版：{old_err}")
            if not new_ok:
                msgs.append(f"新版：{new_err}")
            messagebox.showwarning("跳转失败", "\n".join(msgs))
        else:
            self.log(f"已跳转到 {sheet_name}!{cell_addr}（旧版+新版）")

    def stop_compare(self):
        self.stop_event.set()
        self.stop_btn.configure(state='disabled')
        self.log("正在停止检查...")

if __name__ == "__main__":
    app = tb.Window(themename="flatly")
    viewer = DiffViewer(app)
    app.mainloop()
