"""
Excel 报告检查工具 v3.9 - 最终修复版
包含所有确认功能：常规差异对比、规则检查模式、进度显示、退出规则模式、
规则编辑器左右布局、双击跳转、13项检查、range模式等。
"""
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import threading
import time
import os
import json
import zipfile
import re
import pythoncom
import win32com.client
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from lxml import etree

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

# ======================== 检测项默认配置 ========================
DEFAULT_CHECK_OPTIONS = {
    'value': True,
    'formula': True,
    'rich_text': True,
    'font': True,
    'fill': True,
    'border': True,
    'alignment': True,
    'number_format': True,
    'merged_cells': True,
    'row_height': True,
    'col_width': True,
    'images': True,
    'conditional_format': True,
}

CHECK_OPTION_LABELS = {
    'value': '值变化',
    'formula': '公式变化',
    'rich_text': '富文本',
    'font': '字体',
    'fill': '填充/背景色',
    'border': '边框',
    'alignment': '对齐方式',
    'number_format': '数字格式',
    'merged_cells': '合并单元格',
    'row_height': '行高',
    'col_width': '列宽',
    'images': '图片',
    'conditional_format': '条件格式',
}

CHECK_OPTION_GROUPS = [
    ("内容检测", ['value', 'formula', 'rich_text']),
    ("格式检测", ['font', 'fill', 'border', 'alignment', 'number_format']),
    ("结构检测", ['merged_cells', 'row_height', 'col_width', 'images', 'conditional_format']),
]

# ======================== 优化常量 ========================
PROGRESS_UPDATE_INTERVAL = 200

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

def normalize_color_for_compare(color):
    if color is None:
        return None
    try:
        if hasattr(color, 'type') and color.type == 'auto':
            return None
        if hasattr(color, 'theme') and color.theme is not None:
            tint = color.tint if color.tint is not None else 0.0
            return ('theme', color.theme, round(tint, 4))
        if hasattr(color, 'indexed') and color.indexed is not None:
            return ('indexed', color.indexed)
        if hasattr(color, 'rgb') and color.rgb is not None:
            rgb = color.rgb
            if isinstance(rgb, str):
                if len(rgb) == 8 and rgb[:2] in ('00', 'FF'):
                    return ('rgb', rgb[2:].upper())
                return ('rgb', rgb.upper())
            return ('rgb', str(rgb).upper())
    except Exception:
        pass
    return None

def file_size_mb(path):
    try:
        return os.path.getsize(path) / (1024 * 1024)
    except:
        return 0

def get_sheet_names_fast(file_path):
    if not os.path.isfile(file_path):
        return []
    try:
        zf = zipfile.ZipFile(file_path)
        if 'xl/workbook.xml' in zf.namelist():
            xml_content = zf.read('xl/workbook.xml')
            root = etree.fromstring(xml_content)
            sheets = []
            for sheet in root.findall(f'{NS}sheet'):
                name = sheet.get('name', '')
                sheets.append(name)
            zf.close()
            return sheets
        zf.close()
    except:
        pass
    return []

# ---------------------------- 富文本底层解析 ----------------------------
def _parse_rPr(rPr_elem):
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
    runs = []
    t_elem = si_elem.find(f'{NS}t')
    if t_elem is not None:
        runs.append((t_elem.text or '', None))
        return runs
    for r in si_elem.findall(f'{NS}r'):
        t_elem = r.find(f'{NS}t')
        text = t_elem.text if t_elem is not None else ''
        rPr = r.find(f'{NS}rPr')
        font = _parse_rPr(rPr)
        runs.append((text, font))
    return runs

def parse_rich_text_from_xlsx(xlsx_path):
    result = {}
    if not os.path.isfile(xlsx_path):
        return result
    try:
        zf = zipfile.ZipFile(xlsx_path)
    except:
        return result
    shared_runs = []
    if 'xl/sharedStrings.xml' in zf.namelist():
        ss_xml = zf.read('xl/sharedStrings.xml')
        ss_root = etree.fromstring(ss_xml)
        for si in ss_root.findall(f'{NS}si'):
            runs = _extract_runs_from_si(si)
            shared_runs.append(runs)
    sheet_name_map = {}
    R_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    rid_to_target = {}
    if 'xl/_rels/workbook.xml.rels' in zf.namelist():
        rels_xml = zf.read('xl/_rels/workbook.xml.rels')
        rels_root = etree.fromstring(rels_xml)
        for rel in rels_root.findall(f'{R_NS}Relationship'):
            rid = rel.get('Id', '')
            target = rel.get('Target', '')
            rid_to_target[rid] = target
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
                    filename = target.split('/')[-1]
                    if filename.endswith('.xml'):
                        sheet_name_map[filename] = name
    if not sheet_name_map:
        sheet_files = sorted(
            [n for n in zf.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')],
            key=lambda x: int(x.split('sheet')[1].split('.xml')[0]) if 'sheet' in x else 0
        )
        for idx, f in enumerate(sheet_files, 1):
            sheet_name_map[f.split('/')[-1]] = f'Sheet{idx}'
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
            t = c.get('t', '')
            if t == 's':
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
    if (not runs1 and not runs2):
        return None
    if runs1 is None and runs2 is None:
        return None
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

# ======================== XML级预对比 ========================
def _get_sheet_xml_from_zip(xlsx_path, sheet_filename):
    try:
        zf = zipfile.ZipFile(xlsx_path)
        data = zf.read(f'xl/worksheets/{sheet_filename}')
        zf.close()
        return data
    except:
        return None

def _get_sheet_file_mapping(xlsx_path):
    mapping = {}
    try:
        zf = zipfile.ZipFile(xlsx_path)
    except:
        return mapping
    R_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    rid_to_target = {}
    if 'xl/_rels/workbook.xml.rels' in zf.namelist():
        rels_xml = zf.read('xl/_rels/workbook.xml.rels')
        rels_root = etree.fromstring(rels_xml)
        for rel in rels_root.findall(f'{R_NS}Relationship'):
            rid = rel.get('Id', '')
            target = rel.get('Target', '')
            rid_to_target[rid] = target
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
                    filename = target.split('/')[-1]
                    if filename.endswith('.xml'):
                        mapping[name] = filename
    zf.close()
    return mapping

def xml_fast_compare_sheets(old_path, new_path, common_sheets, log_callback=None):
    results = {}
    old_map = _get_sheet_file_mapping(old_path)
    new_map = _get_sheet_file_mapping(new_path)
    identical_count = 0
    diff_count = 0
    for sheet_name in common_sheets:
        old_fn = old_map.get(sheet_name, '')
        new_fn = new_map.get(sheet_name, '')
        if not old_fn or not new_fn:
            results[sheet_name] = True
            continue
        old_xml = _get_sheet_xml_from_zip(old_path, old_fn)
        new_xml = _get_sheet_xml_from_zip(new_path, new_fn)
        if old_xml is None or new_xml is None:
            results[sheet_name] = True
            continue
        if old_xml == new_xml:
            results[sheet_name] = False
            identical_count += 1
        else:
            results[sheet_name] = True
            diff_count += 1
    if log_callback:
        log_callback(f"[XML预检] {identical_count} 个sheet完全相同(跳过)，{diff_count} 个sheet需要深入对比")
    return results

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
        except Exception as e:
            print(f"加载规则文件失败: {e}")
            return False

    def locate_all(self, workbook):
        results = {}
        for rule in self.rules:
            name = rule.get('name', 'unnamed')
            try:
                result = self.locate(workbook, rule)
                results[name] = result
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
        # 默认在整个工作表查找，允许规则中覆盖 search_in
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

        anchor = self._find_anchor(ws, anchor_cfg)
        if not anchor:
            return {'error': f'Anchor "{anchor_cfg.get("text")}" not found'}

        start_row = anchor[0] + row_offset
        start_col = anchor[1] + col_offset
        if start_row < 1 or start_col < 1:
            return {'error': 'Target out of range'}

        addresses = []
        values = []
        for r in range(start_row, start_row + row_count):
            if r > (ws.max_row or 1):
                break
            for c in range(start_col, start_col + col_count):
                if c > (ws.max_column or 1):
                    break
                cell = ws.cell(row=r, column=c)
                addr = cell_address(c, r)
                addresses.append(addr)
                values.append(cell.value)

        return {
            'address': addresses[0] if addresses else None,
            'addresses': addresses,
            'values': values,
            'range_count': len(addresses)
        }

# ======================== 检查插件框架 ========================
class CheckPlugin:
    name = ""
    description = ""
    def __init__(self, config=None):
        self.config = config or {}
    def check(self, old_data, new_data, context=None):
        raise NotImplementedError

class MeanDeviationPlugin(CheckPlugin):
    name = "mean_deviation"
    description = "检查均值偏差是否超过阈值"
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
                deviation = abs(new_num - old_num) / abs(old_num)
                if deviation > threshold:
                    results.append({
                        'type': '均值偏差告警',
                        'desc': f'偏差 {deviation:.2%} (阈值 {threshold:.0%}): {old_num:.4f} → {new_num:.4f}',
                        'severity': 'warning'
                    })
        except (ValueError, TypeError):
            pass
        return results

class ParamLockPlugin(CheckPlugin):
    name = "param_lock"
    description = "检查关键参数是否保持不变"
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
    description = "检查数值是否在规格范围内"
    def check(self, old_data, new_data, context=None):
        results = []
        lsl = self.config.get('lsl')
        usl = self.config.get('usl')
        new_val = new_data.get('value') if isinstance(new_data, dict) else new_data
        if new_val is None:
            return results
        try:
            num_val = float(new_val)
            if lsl is not None and num_val < float(lsl):
                results.append({
                    'type': '低于下限',
                    'desc': f'值 {num_val:.4f} < LSL {lsl}',
                    'severity': 'error'
                })
            if usl is not None and num_val > float(usl):
                results.append({
                    'type': '超出上限',
                    'desc': f'值 {num_val:.4f} > USL {usl}',
                    'severity': 'error'
                })
        except (ValueError, TypeError):
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
        except Exception as e:
            print(f"加载插件配置失败: {e}")
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

# ======================== 检查项目集数据结构 ========================
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
        if check_options:
            self.check_options.update(check_options)
        self.plugin_manager = plugin_manager
        self.check_project = check_project
        self.xml_skip_map = {}
        self._log_buffer = []
        self._last_gui_update = 0

    def _with_heartbeat(self, label, fn):
        stop_event = threading.Event()
        t0 = time.time()
        def _heartbeat():
            while not stop_event.is_set():
                stop_event.wait(3.0)
                if stop_event.is_set():
                    break
                elapsed = time.time() - t0
                try:
                    self._buf_log(f"⏳ {label}... 已耗时 {elapsed:.0f}s")
                    self._flush_log(force=True)
                except:
                    pass
        hb = threading.Thread(target=_heartbeat, daemon=True)
        self.progress_mode('indeterminate')
        hb.start()
        try:
            result = fn()
        finally:
            stop_event.set()
            hb.join(timeout=2)
            self.progress_mode('determinate')
        return result

    def _load_rich_text(self):
        self.old_rich = parse_rich_text_from_xlsx(self.old_path)
        self.new_rich = parse_rich_text_from_xlsx(self.new_path)
        self._buf_log(f"富文本解析完成：旧版 {sum(len(v) for v in self.old_rich.values())} 个，"
                      f"新版 {sum(len(v) for v in self.new_rich.values())} 个")

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
        if self.mode == 'diff':
            self._run_diff_mode(old_wb, new_wb)
        else:
            self._run_rule_mode(old_wb, new_wb)
        self.progress(95, "生成报告...")
        self._flush_log(force=True)
        total_time = time.time() - start_time
        self.progress(100, "对比完成")
        self._buf_log(f"总耗时: {total_time:.1f}s | 差异: {self.stats['diff_cells']} 处单元格, {len(self.sheet_diffs)} 处Sheet")
        self._flush_log(force=True)
        return True

    def _load_workbooks(self):
        try:
            self.progress(5, "正在加载旧版文件...")
            self._flush_log(force=True)
            old_wb = load_workbook(self.old_path, data_only=False)
            self._buf_log(f"旧版加载完成: {len(old_wb.sheetnames)} 个sheet")

            self.progress(15, "正在加载新版文件...")
            self._flush_log(force=True)
            new_wb = load_workbook(self.new_path, data_only=False)
            self._buf_log(f"新版加载完成: {len(new_wb.sheetnames)} 个sheet")
            self._flush_log(force=True)

            return old_wb, new_wb
        except Exception as e:
            self._buf_log(f"加载工作簿失败: {e}")
            return None, None

    def _run_diff_mode(self, old_wb, new_wb):
        self._compare_sheets(old_wb, new_wb)
        total = len(old_wb.sheetnames)
        for idx, sheet_name in enumerate(old_wb.sheetnames, 1):
            if self.stop_event.is_set():
                self._buf_log(f"用户请求停止，已跳过剩余 {total - idx + 1} 个sheet")
                self._flush_log(force=True)
                raise KeyboardInterrupt
            pct = 25 + int(55 * idx / total)
            self.progress(pct, f"对比 {sheet_name}... ({idx}/{total})")
            self._flush_log(force=True)
            if sheet_name in new_wb.sheetnames:
                old_ws = old_wb[sheet_name]
                new_ws = new_wb[sheet_name]
                self._compare_worksheet(old_ws, new_ws, sheet_name)
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

    def _run_rule_mode(self, old_wb, new_wb):
        if self.check_project:
            total_rules = len(self.check_project.rules)
            self._buf_log(f"正在执行检查项目规则：共 {total_rules} 条规则")
            self._flush_log(force=True)
            for idx, rule in enumerate(self.check_project.rules):
                if self.stop_event.is_set():
                    self._buf_log("用户请求停止，已跳过剩余规则")
                    self._flush_log(force=True)
                    raise KeyboardInterrupt
                pct = 30 + int(60 * (idx / max(total_rules, 1)))
                self.progress(pct, f"正在执行规则: {rule.rule_name}")
                self._buf_log(f"正在执行规则: {rule.rule_name}")
                self._flush_log(force=True)
                self._run_single_rule(rule, old_wb, new_wb)
            self.progress(95, "规则执行完成")
            self._flush_log(force=True)
        else:
            self._buf_log("未加载检查项目，无法执行规则检查")
            self.progress(95, "规则执行完成")
            self._flush_log(force=True)

    def _run_single_rule(self, rule, old_wb, new_wb):
        if not rule.data_source:
            return
        old_locator = DataLocator()
        old_locator.rules = [rule.data_source]
        old_data_list = old_locator.locate_all(old_wb)
        rule_name = rule.data_source.get('name', rule.rule_name)
        old_data = old_data_list.get(rule_name)
        if old_data is None:
            old_data = old_data_list.get('', None)
        new_locator = DataLocator()
        new_locator.rules = [rule.data_source]
        new_data_list = new_locator.locate_all(new_wb)
        new_data = new_data_list.get(rule_name)
        if new_data is None:
            new_data = new_data_list.get('', None)
        if old_data is None or new_data is None:
            self._buf_log(f"规则 [{rule.rule_name}] 数据源定位失败，跳过")
            return

        # 检查 DataLocator 返回的错误信息
        if isinstance(old_data, dict) and 'error' in old_data:
            self._buf_log(f"规则 [{rule.rule_name}] 旧版数据源错误: {old_data['error']}")
            return
        if isinstance(new_data, dict) and 'error' in new_data:
            self._buf_log(f"规则 [{rule.rule_name}] 新版数据源错误: {new_data['error']}")
            return

        # 修复：优先获取 addresses 列表（range 模式）
        address = None
        if isinstance(old_data, dict) and 'addresses' in old_data:
            address = old_data['addresses']
        elif isinstance(old_data, dict) and 'address' in old_data:
            address = old_data['address']
        if not address:
            self._buf_log(f"规则 [{rule.rule_name}] 无法获取地址，跳过")
            return

        sheet_name = rule.data_source.get('sheet', '')
        if sheet_name not in old_wb.sheetnames or sheet_name not in new_wb.sheetnames:
            self._buf_log(f"规则 [{rule.rule_name}] sheet不存在，跳过")
            return

        old_ws = old_wb[sheet_name]
        new_ws = new_wb[sheet_name]

        if isinstance(address, list):
            for addr in address:
                self._check_single_cell(rule, old_ws, new_ws, addr, sheet_name)
        else:
            self._check_single_cell(rule, old_ws, new_ws, address, sheet_name)

    def _check_single_cell(self, rule, old_ws, new_ws, addr, sheet_name):
        col_str = ''.join(ch for ch in addr if ch.isalpha())
        row_str = ''.join(ch for ch in addr if ch.isdigit())
        if not col_str or not row_str:
            return
        col_idx = column_index_from_string(col_str)
        row_idx = int(row_str)
        old_cell = old_ws.cell(row=row_idx, column=col_idx)
        new_cell = new_ws.cell(row=row_idx, column=col_idx)
        for check in rule.checks:
            if not check.enabled:
                continue
            diff = self._compare_by_check_type(check.check_type, old_cell, new_cell, check.options,
                                               old_ws, new_ws, addr, sheet_name)
            if check.expect == 'same' and diff is not None:
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': addr,
                    'type': f"规则检查-{check.check_type}",
                    'desc': diff,
                    'severity': 'error'
                })
                self.stats['diff_cells'] += 1  # 修复：更新统计
            elif check.expect == 'different' and diff is None:
                self.diffs.append({
                    'sheet': sheet_name,
                    'address': addr,
                    'type': f"规则检查-{check.check_type}",
                    'desc': "预期存在差异，但实际相同",
                    'severity': 'warning'
                })

    def _compare_by_check_type(self, check_type, old_cell, new_cell, options=None,
                               old_ws=None, new_ws=None, address=None, sheet_name=''):
        if check_type == 'value':
            old_val = old_cell.value
            new_val = new_cell.value
            if old_val != new_val:
                return f"{old_val} → {new_val}"
            return None
        elif check_type == 'formula':
            old_formula = old_cell.value if isinstance(old_cell.value, str) and old_cell.value.startswith('=') else None
            new_formula = new_cell.value if isinstance(new_cell.value, str) and new_cell.value.startswith('=') else None
            if old_formula != new_formula:
                return f"公式: {old_formula} → {new_formula}"
            return None
        elif check_type == 'rich_text':
            if old_ws is not None and new_ws is not None:
                old_rich = self.old_rich.get(sheet_name, {}).get(address)
                new_rich = self.new_rich.get(sheet_name, {}).get(address)
                return compare_rich_text_runs(old_rich, new_rich)
            return None
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
            if nf1 != nf2:
                return f"{nf1} → {nf2}"
            return None
        elif check_type == 'merged_cells':
            old_merged = str(old_cell.coordinate) if any(str(old_cell.coordinate) in str(m) for m in old_ws.merged_cells.ranges) else None
            new_merged = str(new_cell.coordinate) if any(str(new_cell.coordinate) in str(m) for m in new_ws.merged_cells.ranges) else None
            if old_merged != new_merged:
                return f"合并区域: {old_merged} → {new_merged}"
            return None
        elif check_type == 'row_height':
            old_height = old_ws.row_dimensions[old_cell.row].height if old_cell.row in old_ws.row_dimensions else None
            new_height = new_ws.row_dimensions[new_cell.row].height if new_cell.row in new_ws.row_dimensions else None
            if old_height != new_height:
                return f"行高: {old_height} → {new_height}"
            return None
        elif check_type == 'col_width':
            col_letter = get_column_letter(old_cell.column)
            old_width = old_ws.column_dimensions[col_letter].width if col_letter in old_ws.column_dimensions else None
            new_width = new_ws.column_dimensions[col_letter].width if col_letter in new_ws.column_dimensions else None
            if old_width != new_width:
                return f"列宽({col_letter}): {old_width} → {new_width}"
            return None
        elif check_type == 'images':
            old_imgs = self._get_images_from_ws(old_ws)
            new_imgs = self._get_images_from_ws(new_ws)
            old_has = any(addr == address for addr, _, _ in old_imgs)
            new_has = any(addr == address for addr, _, _ in new_imgs)
            if old_has != new_has:
                return f"图片存在: {old_has} → {new_has}"
            return None
        elif check_type == 'conditional_format':
            old_cf = self._get_conditional_format_for_cell(old_ws, address)
            new_cf = self._get_conditional_format_for_cell(new_ws, address)
            if old_cf != new_cf:
                return f"条件格式: {old_cf} → {new_cf}"
            return None
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
        if self.sheet_diffs:
            self._buf_log(f"Sheet结构: {len([s for s in self.sheet_diffs if s['type']=='新增'])} 新增, "
                          f"{len([s for s in self.sheet_diffs if s['type']=='删除'])} 删除")

    def _row_has_data(self, ws, row_idx, max_col_check):
        for c in range(1, max_col_check + 1):
            cell = ws.cell(row=row_idx, column=c)
            if cell.value is not None:
                return True
        return False

    def _col_has_data(self, ws, col_idx, max_row_check):
        for r in range(1, max_row_check + 1):
            cell = ws.cell(row=r, column=col_idx)
            if cell.value is not None:
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
        self._buf_log(f" [范围裁切] openpyxl报告 {max_row}行x{max_col}列 → 实际 {real_max_row}行x{real_max_col}列")
        return 1, 1, real_max_row, real_max_col

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

        batch_size = PROGRESS_UPDATE_INTERVAL
        row_count = 0
        for row_idx in range(1, max_row + 1):
            if self.stop_event.is_set():
                raise KeyboardInterrupt
            old_row_data = None
            new_row_data = None
            if row_idx <= old_max_row:
                old_row_data = list(old_ws.iter_rows(min_row=row_idx, max_row=row_idx, min_col=1, max_col=max_col, values_only=False))
                old_row_data = old_row_data[0] if old_row_data else []
            if row_idx <= new_max_row:
                new_row_data = list(new_ws.iter_rows(min_row=row_idx, max_row=row_idx, min_col=1, max_col=max_col, values_only=False))
                new_row_data = new_row_data[0] if new_row_data else []
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
                # 值检查（如果启用）
                if check_value:
                    val_diff = self._get_cell_diff(old_cell, new_cell)
                    if val_diff:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '内容变化', 'desc': val_diff})
                        self.stats['diff_cells'] += 1
                        # 值不同时，后续其他格式检查仍然执行，但跳过公式独立检查？
                        # 原逻辑是continue，表示如果值有差异，跳过格式检查？原代码continue，但可能导致格式差异丢失。
                        # 我们保持原逻辑，不改变（原代码是continue，跳过后续格式检查）。如果觉得不合理，可调整。
                        # 这里我们保留原行为，因为用户未提出修改。
                        continue
                # 修复：公式独立检查（当值检查未启用时）
                if check_formula and not check_value:
                    old_formula = old_cell.value if isinstance(old_cell.value, str) and old_cell.value.startswith('=') else None
                    new_formula = new_cell.value if isinstance(new_cell.value, str) and new_cell.value.startswith('=') else None
                    if old_formula != new_formula:
                        diff = f"公式: {old_formula} → {new_formula}"
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '公式变化', 'desc': diff})
                        self.stats['diff_cells'] += 1
                        continue
                # 格式检查（如果值检查通过或未启用）
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
                    if nf1 != nf2:
                        self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '数字格式变化', 'desc': f'{nf1} → {nf2}'})
                row_count += 1
                if row_count % batch_size == 0:
                    cur_pct = 25 + int(55 * (row_idx / max_row))
                    self.progress(cur_pct, f" {sheet_name}: {row_idx}/{max_row}行...")
                    self._flush_log(force=True)

        # 富文本对比
        if opts.get('rich_text', True):
            old_sheet_rich = self.old_rich.get(sheet_name, {})
            new_sheet_rich = self.new_rich.get(sheet_name, {})
            all_refs = set(old_sheet_rich.keys()) | set(new_sheet_rich.keys())
            for ref in all_refs:
                r1 = old_sheet_rich.get(ref)
                r2 = new_sheet_rich.get(ref)
                rt_diff = compare_rich_text_runs(r1, r2)
                if rt_diff:
                    self.diffs.append({'sheet': sheet_name, 'address': ref, 'type': '富文本变化', 'desc': rt_diff})

        # 合并单元格
        if opts.get('merged_cells', True):
            self._compare_merged_cells(old_ws, new_ws, sheet_name)
        # 行列维度
        if opts.get('row_height', True) or opts.get('col_width', True):
            self._compare_row_col_dimensions(old_ws, new_ws, sheet_name)
        # 图片
        if opts.get('images', True):
            self._compare_images(old_ws, new_ws, sheet_name)
        # 条件格式
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
            s1 = str(v1) if v1 is not None else ''
            s2 = str(v2) if v2 is not None else ''
            max_len = 120
            if len(s1) > max_len:
                s1 = s1[:max_len] + '...'
            if len(s2) > max_len:
                s2 = s2[:max_len] + '...'
            return f"{s1} → {s2}"
        return None

    def _cmp_font(self, f1, f2):
        changes = []
        n1 = f1.name if f1.name is not None else ''
        n2 = f2.name if f2.name is not None else ''
        if n1 != n2:
            changes.append(f"字体: {n1}→{n2}")
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
            return (f"类型: {f1.fill_type}→{f2.fill_type}, "
                    f"颜色: {rgb_to_hex(f1.start_color)}→{rgb_to_hex(f2.start_color)}")
        return None

    def _cmp_border(self, b1, b2):
        parts = []
        for side in ['left', 'right', 'top', 'bottom']:
            s1 = getattr(b1, side)
            s2 = getattr(b2, side)
            st1 = s1.style if s1.style is not None else None
            st2 = s2.style if s2.style is not None else None
            c1 = normalize_color_for_compare(s1.color)
            c2 = normalize_color_for_compare(s2.color)
            if st1 != st2 or c1 != c2:
                parts.append(f"{side}: {s1.style}/{rgb_to_hex(s1.color)}→{s2.style}/{rgb_to_hex(s2.color)}")
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

    def _compare_merged_cells(self, old_ws, new_ws, sheet_name):
        old_merged = set(str(m) for m in old_ws.merged_cells.ranges)
        new_merged = set(str(m) for m in new_ws.merged_cells.ranges)
        for addr in new_merged - old_merged:
            self.diffs.append({'sheet': sheet_name, 'address': addr.split(':')[0], 'type': '合并新增', 'desc': f'新增合并区域 {addr}'})
        for addr in old_merged - new_merged:
            self.diffs.append({'sheet': sheet_name, 'address': addr.split(':')[0], 'type': '合并删除', 'desc': f'删除合并区域 {addr}'})

    def _compare_row_col_dimensions(self, old_ws, new_ws, sheet_name):
        opts = self.check_options
        if opts.get('row_height', True):
            all_rows = set(old_ws.row_dimensions.keys()) | set(new_ws.row_dimensions.keys())
            for row_idx in all_rows:
                oh = old_ws.row_dimensions[row_idx].height if row_idx in old_ws.row_dimensions else None
                nh = new_ws.row_dimensions[row_idx].height if row_idx in new_ws.row_dimensions else None
                if oh != nh:
                    self.diffs.append({'sheet': sheet_name, 'address': f"A{row_idx}", 'type': '行高变化', 'desc': f'行高: {oh} → {nh}'})
        if opts.get('col_width', True):
            all_cols = set(old_ws.column_dimensions.keys()) | set(new_ws.column_dimensions.keys())
            for col_letter in all_cols:
                ow = old_ws.column_dimensions[col_letter].width if col_letter in old_ws.column_dimensions else None
                nw = new_ws.column_dimensions[col_letter].width if col_letter in new_ws.column_dimensions else None
                if ow != nw:
                    col_idx = column_index_from_string(col_letter)
                    addr = cell_address(col_idx, 1)
                    self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '列宽变化', 'desc': f'列宽({col_letter}): {ow} → {nw}'})

    def _get_images_from_ws(self, ws):
        images = []
        if hasattr(ws, '_images') and ws._images:
            for img in ws._images:
                try:
                    anchor = img.anchor
                    if anchor and hasattr(anchor, '_from'):
                        col = anchor._from.col + 1
                        row = anchor._from.row + 1
                        addr = cell_address(col, row)
                        images.append((addr, img.width, img.height))
                except Exception as e:
                    self._buf_log(f"图片提取异常: {e}")
            if images:
                return sorted(images, key=lambda x: x[0])
        if hasattr(ws, '_drawing') and ws._drawing:
            for anchor in ws._drawing.anchors:
                if hasattr(anchor, 'image'):
                    img = anchor.image
                    col = (anchor._from.col if hasattr(anchor, '_from') else 0) + 1
                    row = (anchor._from.row if hasattr(anchor, '_from') else 0) + 1
                    addr = cell_address(col, row)
                    images.append((addr, img.width, img.height))
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
            w_rel = abs(w1 - w2) / max(w1, w2) if max(w1, w2) > 0 else 0
            h_rel = abs(h1 - h2) / max(h1, h2) if max(h1, h2) > 0 else 0
            w_abs = abs(w1 - w2)
            h_abs = abs(h1 - h2)
            if w_rel > 0.01 or h_rel > 0.01:
                if w_abs > 1 or h_abs > 1:
                    changed.append((addr, w1, h1, w2, h2))
        diff_count = len(added) + len(removed) + len(changed)
        if diff_count > 0:
            self.stats['images_diff'] += diff_count
            for addr in sorted(added):
                w, h = new_set[addr]
                self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '图片新增', 'desc': f'新增图片 ({w:.0f}x{h:.0f})'})
            for addr in sorted(removed):
                w, h = old_set[addr]
                self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '图片删除', 'desc': f'删除图片 ({w:.0f}x{h:.0f})'})
            for addr, w1, h1, w2, h2 in sorted(changed, key=lambda x: x[0]):
                self.diffs.append({'sheet': sheet_name, 'address': addr, 'type': '图片尺寸变化', 'desc': f'图片尺寸: {w1:.0f}x{h1:.0f} → {w2:.0f}x{h2:.0f}'})

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
                self.diffs.append({'sheet': sheet_name, 'address': start, 'type': '条件格式新增', 'desc': f'新增条件格式范围: {rng}'})
            elif new_cf is None:
                start = rng.split(':')[0] if ':' in rng else rng
                self.diffs.append({'sheet': sheet_name, 'address': start, 'type': '条件格式删除', 'desc': f'删除条件格式范围: {rng}'})
            else:
                old_rules = [(r.type, r.priority, str(r.formula), str(r.dxf)) for r in old_cf.rules]
                new_rules = [(r.type, r.priority, str(r.formula), str(r.dxf)) for r in new_cf.rules]
                if old_rules != new_rules:
                    start = rng.split(':')[0] if ':' in rng else rng
                    self.diffs.append({'sheet': sheet_name, 'address': start, 'type': '条件格式修改', 'desc': f'条件格式规则变化，范围: {rng}'})

# ---------------------------- 检测项设置弹窗 ----------------------------
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

        for group_name, keys in CHECK_OPTION_GROUPS:
            lf = tb.Labelframe(main_frame, text=group_name, padding=(8, 5))
            lf.pack(fill='x', pady=(0, 8))
            row_frame = None
            for i, key in enumerate(keys):
                if i % 3 == 0:
                    row_frame = tb.Frame(lf)
                    row_frame.pack(fill='x', pady=1)
                var = tk.BooleanVar(value=current_options.get(key, True))
                self.vars[key] = var
                cb = tb.Checkbutton(row_frame, text=CHECK_OPTION_LABELS[key], variable=var, bootstyle="round-toggle")
                cb.pack(side='left', padx=(0, 15))

        btn_frame = tb.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(12, 0))
        tb.Button(btn_frame, text="全选", width=8, command=self._select_all).pack(side='left', padx=(0, 5))
        tb.Button(btn_frame, text="全不选", width=8, command=self._deselect_all).pack(side='left', padx=(0, 5))
        tb.Button(btn_frame, text="取消", width=8, command=self._on_cancel).pack(side='right', padx=(5, 0))
        tb.Button(btn_frame, text="确定", bootstyle=PRIMARY, width=8, command=self._on_ok).pack(side='right')

        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        self.geometry(f"+{pw - w//2}+{ph - h//2}")
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

# ======================== 检查项目集设置窗口 ========================
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
        tb.Label(info_frame, text="项目名称:").grid(row=0, column=0, sticky='w', padx=(0,5))
        self.project_name_var = tk.StringVar(value=self.project.project_name)
        tb.Entry(info_frame, textvariable=self.project_name_var, width=30).grid(row=0, column=1, sticky='w')
        tb.Label(info_frame, text="版本:").grid(row=0, column=2, sticky='w', padx=(15,5))
        self.version_var = tk.StringVar(value=self.project.version)
        tb.Entry(info_frame, textvariable=self.version_var, width=10).grid(row=0, column=3, sticky='w')
        tb.Label(info_frame, text="描述:").grid(row=0, column=4, sticky='w', padx=(15,5))
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
            ds_str = f"{ds.get('sheet','?')} | {ds.get('anchor',{}).get('text','?')}"
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
        confirm = messagebox.askyesno("确认", "确定删除该规则？")
        if confirm:
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
        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile=f"{self.project.project_name}.json"
        )
        if filepath:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(self.project.to_dict(), f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", f"项目已保存到: {filepath}")

    def load_project(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filepath:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.project = CheckProject.from_dict(data)
            self.project_name_var.set(self.project.project_name)
            self.version_var.set(self.project.version)
            self.desc_var.set(self.project.description)
            self._refresh_rule_list()
            messagebox.showinfo("成功", "项目已加载")

    def apply_project(self):
        self.project.project_name = self.project_name_var.get()
        self.project.version = self.version_var.get()
        self.project.description = self.desc_var.get()
        self.result = self.project
        self.destroy()

# ======================== 规则编辑器对话框 ========================
class RuleEditorDialog(tb.Toplevel):
    def __init__(self, parent, old_path, new_path, rule=None):
        super().__init__(parent)
        self.title("编辑规则")
        self.geometry("1100x700")
        self.parent = parent
        self.old_path = old_path
        self.new_path = new_path
        self.result = None
        self.rule = rule if rule else CheckRule()

        self._build_ui()
        self._load_rule_data()
        self.wait_window()

    def _get_sheet_names(self):
        paths = [self.old_path, self.new_path]
        for p in paths:
            if os.path.isfile(p):
                sheets = get_sheet_names_fast(p)
                if sheets:
                    return sheets
        return []

    def _build_ui(self):
        main_frame = tb.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # 左侧：数据源配置（加宽，宽度420）
        left_frame = tb.Frame(main_frame, width=420)
        left_frame.pack(side='left', fill='y', padx=(0,10))
        left_frame.pack_propagate(False)

        ds_frame = tb.Labelframe(left_frame, text="数据源配置", padding=10)
        ds_frame.pack(fill='both', expand=True, pady=5)

        # 所有字段竖向排列
        # 规则名称
        tb.Label(ds_frame, text="规则名称:").pack(anchor='w')
        self.rule_name_var = tk.StringVar()
        tb.Entry(ds_frame, textvariable=self.rule_name_var, width=40).pack(fill='x', pady=2)

        # Sheet
        tb.Label(ds_frame, text="Sheet:").pack(anchor='w')
        self.sheet_var = tk.StringVar()
        sheets = self._get_sheet_names()
        self.sheet_cb = tb.Combobox(ds_frame, textvariable=self.sheet_var, values=sheets, width=38)
        self.sheet_cb.pack(fill='x', pady=2)

        # 锚点文字
        tb.Label(ds_frame, text="锚点文字:").pack(anchor='w')
        self.anchor_text_var = tk.StringVar()
        tb.Entry(ds_frame, textvariable=self.anchor_text_var, width=40).pack(fill='x', pady=2)

        # 搜索范围
        tb.Label(ds_frame, text="搜索范围:").pack(anchor='w')
        self.search_in_var = tk.StringVar(value='all')
        tb.Combobox(ds_frame, textvariable=self.search_in_var, values=['all','first_row','first_col'], width=38).pack(fill='x', pady=2)

        # 模式
        tb.Label(ds_frame, text="模式:").pack(anchor='w')
        self.mode_var = tk.StringVar(value='offset')
        self.mode_cb = tb.Combobox(ds_frame, textvariable=self.mode_var, values=['offset','collect','intersection','range'], width=38)
        self.mode_cb.pack(fill='x', pady=2)
        self.mode_cb.bind('<<ComboboxSelected>>', lambda e: self._build_param_fields())

        # 动态参数区（竖向排列）
        self.param_frame = tb.Frame(ds_frame)
        self.param_frame.pack(fill='x', pady=2)
        self._build_param_fields()

        # 按钮在左侧底部
        btn_frame = tb.Frame(left_frame)
        btn_frame.pack(side='bottom', fill='x', pady=5)
        tb.Button(btn_frame, text="确定", bootstyle=PRIMARY, width=8, command=self.on_ok).pack(side='left', padx=5)
        tb.Button(btn_frame, text="取消", bootstyle="outline", width=8, command=self.on_cancel).pack(side='right', padx=5)

        # 右侧：检查项配置（使用grid布局）
        right_frame = tb.Frame(main_frame)
        right_frame.pack(side='right', fill='both', expand=True)

        check_frame = tb.Labelframe(right_frame, text="检查项（可多选）", padding=10)
        check_frame.pack(fill='both', expand=True)

        # 可滚动区域
        canvas = tk.Canvas(check_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(check_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tb.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.check_vars = {}
        self.expect_vars = {}
        for group_name, keys in CHECK_OPTION_GROUPS:
            lf = tb.Labelframe(scrollable_frame, text=group_name, padding=(8,5))
            lf.pack(fill='x', pady=(0,8), padx=5)
            for key in keys:
                row = tb.Frame(lf)
                row.pack(fill='x', pady=1)
                # 左列：开关 + 名称
                left_cell = tb.Frame(row)
                left_cell.pack(side='left', fill='x', expand=True)
                var = tk.BooleanVar(value=True)
                self.check_vars[key] = var
                cb = tb.Checkbutton(left_cell, text=CHECK_OPTION_LABELS[key], variable=var, bootstyle="round-toggle")
                cb.pack(side='left', anchor='w')
                # 右列：期望 + 下拉
                right_cell = tb.Frame(row)
                right_cell.pack(side='right')
                tb.Label(right_cell, text="期望:").pack(side='left', padx=(10,2))
                expect_var = tk.StringVar(value='same')
                tb.Combobox(right_cell, textvariable=expect_var, values=['same','different'], width=8).pack(side='left')
                self.expect_vars[key] = expect_var

    def _build_param_fields(self):
        # 清空旧内容
        for widget in self.param_frame.winfo_children():
            widget.destroy()
        mode = self.mode_var.get()
        if mode == 'offset':
            # 行偏移
            tb.Label(self.param_frame, text="行偏移:").pack(anchor='w')
            self.offset_row_var = tk.StringVar(value='0')
            tb.Entry(self.param_frame, textvariable=self.offset_row_var, width=10).pack(fill='x', pady=2)
            # 列偏移
            tb.Label(self.param_frame, text="列偏移:").pack(anchor='w')
            self.offset_col_var = tk.StringVar(value='0')
            tb.Entry(self.param_frame, textvariable=self.offset_col_var, width=10).pack(fill='x', pady=2)
        elif mode == 'collect':
            tb.Label(self.param_frame, text="方向:").pack(anchor='w')
            self.collect_dir_var = tk.StringVar(value='down')
            tb.Combobox(self.param_frame, textvariable=self.collect_dir_var, values=['down','right'], width=10).pack(fill='x', pady=2)
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

    def _load_rule_data(self):
        self.rule_name_var.set(self.rule.rule_name)
        ds = self.rule.data_source
        self.sheet_var.set(ds.get('sheet', ''))
        self.anchor_text_var.set(ds.get('anchor', {}).get('text', ''))
        self.search_in_var.set(ds.get('search_in', 'all'))

        # 先设置模式，再构建参数区，确保对应变量已创建
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
        elif mode == 'intersection':
            row_anchor = ds.get('row_anchor', {})
            col_anchor = ds.get('col_anchor', {})
            self.row_anchor_text_var.set(row_anchor.get('text', '').strip())
            self.col_anchor_text_var.set(col_anchor.get('text', '').strip())

        # 加载检查项
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
                'col_count': int(self.range_col_count_var.get())
            }
        self.rule.data_source = ds

        checks = []
        for key, var in self.check_vars.items():
            if var.get():
                check_item = CheckItemConfig(
                    check_type=key,
                    enabled=True,
                    expect=self.expect_vars[key].get()
                )
                checks.append(check_item)
        self.rule.checks = checks
        self.result = self.rule
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

# ======================== GUI ========================
class DiffViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel报告检查工具")  # 工具名称修改
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

        tree_frame = tb.Frame(root, padding=(5, 0))
        tree_frame.pack(fill='both', expand=True)

        self.tree = tb.Treeview(tree_frame, columns=('address', 'type'), show='tree headings', bootstyle=PRIMARY)
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
        self.tree.bind('<Double-1>', self.on_tree_double_click)

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
        # 修复：只允许 .xlsx 文件
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
        else:
            pass

    def clear_check_project(self):
        """退出规则模式，恢复常规差异对比，并清空结果"""
        self.check_project = None
        self.start_btn.configure(text="常规差异对比", width=12)
        self.settings_btn.configure(state='normal')
        self.config_btn.configure(text="导入规则", command=self.load_check_project)
        self.project_btn.configure(text="进阶检查规则", command=self.open_project_dialog)
        # 清空结果
        self.tree.delete(*self.tree.get_children())
        self.detail.delete('1.0', 'end')
        self.diff_items = []
        self.result_data = None
        self.log("已退出规则检查模式，恢复常规差异对比")

    def load_check_project(self):
        """加载检查项目集JSON文件，自动切换规则模式"""
        filepath = filedialog.askopenfilename(
            title="选择检查项目集文件",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not filepath:
            return
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.check_project = CheckProject.from_dict(data)
            # 更新UI
            self.start_btn.configure(text=f"规则检查\n（{self.check_project.project_name}）", width=16)
            self.settings_btn.configure(state='disabled')
            self.config_btn.configure(text="退出规则模式", command=self.clear_check_project)
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
                self.progress.start(15)
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
        # 修复：检查扩展名
        if not old.lower().endswith('.xlsx') or not new.lower().endswith('.xlsx'):
            messagebox.showerror("错误", "仅支持 .xlsx 格式文件")
            return
        self.stop_event.clear()
        self.start_btn.configure(state='disabled')
        self.stop_btn.configure(state='normal')
        self.tree.delete(*self.tree.get_children())
        self.detail.delete('1.0', 'end')
        self.diff_items = []

        if self.check_project:
            def worker():
                comparer = None
                try:
                    comparer = OpenpyxlComparer(old, new, self.log, self.update_progress,
                                                check_options=None, plugin_manager=None,
                                                progress_mode_fn=self.set_progress_mode,
                                                check_project=self.check_project,
                                                stop_event=self.stop_event,
                                                mode='rule')
                    comparer.run()
                    self.result_data = (comparer.diffs, comparer.sheet_diffs, comparer.stats)
                    self.root.after(0, self.populate_tree)
                except KeyboardInterrupt:
                    if comparer:
                        self.result_data = (comparer.diffs, comparer.sheet_diffs, comparer.stats)
                    else:
                        self.result_data = ([], [], {'total_cells':0, 'diff_cells':0, 'added_sheets':[], 'removed_sheets':[], 'images_diff':0})
                    self.root.after(0, self.populate_tree)
                except Exception as e:
                    import traceback
                    tb_str = traceback.format_exc()
                    error_msg = f"{str(e)}\n\n{tb_str}"
                    self.root.after(0, lambda msg=error_msg: messagebox.showerror("对比失败", msg))
                finally:
                    self.root.after(0, self.on_comparison_finished)
            threading.Thread(target=worker, daemon=True).start()
        else:
            current_opts = dict(self.check_options)
            pm = self.plugin_manager
            def worker():
                comparer = None
                try:
                    comparer = OpenpyxlComparer(old, new, self.log, self.update_progress,
                                                check_options=current_opts, plugin_manager=pm,
                                                progress_mode_fn=self.set_progress_mode,
                                                check_project=None,
                                                stop_event=self.stop_event,
                                                mode='diff')
                    comparer.run()
                    self.result_data = (comparer.diffs, comparer.sheet_diffs, comparer.stats)
                    self.root.after(0, self.populate_tree)
                except KeyboardInterrupt:
                    if comparer:
                        self.result_data = (comparer.diffs, comparer.sheet_diffs, comparer.stats)
                    else:
                        self.result_data = ([], [], {'total_cells':0, 'diff_cells':0, 'added_sheets':[], 'removed_sheets':[], 'images_diff':0})
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
        if sheet_diffs:
            sn = self.tree.insert('', 'end', text='📋 Sheet 结构差异', open=True)
            for sd in sheet_diffs:
                node = self.tree.insert(sn, 'end', text=sd['desc'], values=(sd['name'], sd['type']))
                self.diff_items.append((node, {'type': 'sheet_struct', 'data': sd}))
        plugin_diffs = [d for d in diffs if d['sheet'] == '🔍 数据检查']
        normal_diffs = [d for d in diffs if d['sheet'] != '🔍 数据检查']
        if plugin_diffs:
            pn = self.tree.insert('', 'end', text='🔍 数据检查结果', open=True)
            for d in plugin_diffs:
                node = self.tree.insert(pn, 'end', text=d['desc'][:80], values=(d['address'], d['type']))
                self.diff_items.append((node, {'type': 'cell', 'data': d}))
        dmap = {}
        for d in normal_diffs:
            dmap.setdefault(d['sheet'], []).append(d)
        for sname, items in sorted(dmap.items()):
            pn = self.tree.insert('', 'end', text=f"📄 {sname}", open=True)
            for d in items:
                node = self.tree.insert(pn, 'end', text=d['desc'][:80], values=(d['address'], d['type']))
                self.diff_items.append((node, {'type': 'cell', 'data': d}))
        enabled = sum(1 for v in self.check_options.values() if v)
        total = len(self.check_options)
        plugin_info = f"，{len(plugin_diffs)} 个插件告警" if plugin_diffs else ""
        self.log(f"树形列表已加载，单击查看详情，双击跳转（启用 {enabled}/{total} 项检测{plugin_info}）")

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
        sel = self.tree.selection()
        if not sel:
            return
        node = sel[0]
        item_info = next((d for n, d in self.diff_items if n == node), None)
        if not item_info:
            return
        if item_info['type'] != 'cell':
            return
        data = item_info['data']
        sheet_name = data.get('sheet')
        cell_addr = data.get('address')
        if not sheet_name or not cell_addr:
            return
        # 修复：地址格式验证
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
