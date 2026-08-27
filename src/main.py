""" Excel 差异对比工具 v3.1 - ttkbootstrap flatly 主题版 - v3.0: 13项检测开关、公式/字体/填充Bug修复 - v3.1: 提速优化、语义定位引擎、检查插件框架 """
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import threading
import time
import os
import json
import zipfile
import pythoncom
import win32com.client
from win32com.client import constants
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from lxml import etree
from itertools import zip_longest

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

# ======================== 检测项默认配置 ========================
DEFAULT_CHECK_OPTIONS = {
    'value': True, 'formula': True, 'rich_text': True,
    'font': True, 'fill': True, 'border': True, 'alignment': True, 'number_format': True,
    'merged_cells': True, 'row_height': True, 'col_width': True, 'images': True, 'conditional_format': True,
}

CHECK_OPTION_LABELS = {
    'value': '值变化', 'formula': '公式变化', 'rich_text': '富文本',
    'font': '字体', 'fill': '填充/背景色', 'border': '边框', 'alignment': '对齐方式', 'number_format': '数字格式',
    'merged_cells': '合并单元格', 'row_height': '行高', 'col_width': '列宽', 'images': '图片', 'conditional_format': '条件格式',
}

CHECK_OPTION_GROUPS = [
    ("内容检测", ['value', 'formula', 'rich_text']),
    ("格式检测", ['font', 'fill', 'border', 'alignment', 'number_format']),
    ("结构检测", ['merged_cells', 'row_height', 'col_width', 'images', 'conditional_format']),
]

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

# ---------------------------- v3.1 语义定位引擎 ----------------------------
class DataLocator:
    """ 基于规则的数据定位引擎。 支持三种模式： - offset: 锚点+固定偏移，获取单个值 - collect: 锚点+方向收集，获取一组数据 - intersection: 行列交叉定位，获取交叉点值 """
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
        """对所有规则执行定位，返回 {rule_name: located_data}"""
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
        return {'error': f'Unknown mode: {mode}'}
    
    def _find_anchor(self, ws, anchor_cfg):
        """查找锚点单元格，返回(row, col)或None"""
        text = anchor_cfg.get('text', '')
        search_in = anchor_cfg.get('search_in', 'all')  # first_row, first_col, all
        
        if search_in == 'first_row':
            for col in range(1, (ws.max_column or 1) + 1):
                val = ws.cell(1, col).value
                if val is not None and text == str(val).strip():
                    return (1, col)
        elif search_in == 'first_col':
            for row in range(1, (ws.max_row or 1) + 1):
                val = ws.cell(row, 1).value
                if val is not None and text == str(val).strip():
                    return (row, 1)
        else:  # 'all'
            for row in range(1, (ws.max_row or 1) + 1):
                for col in range(1, (ws.max_column or 1) + 1):
                    val = ws.cell(row, col).value
                    if val is not None and text == str(val).strip():
                        return (row, col)
        return None
    
    def _locate_offset(self, ws, rule):
        """锚点+偏移模式"""
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
        """锚点+方向收集模式"""
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
        """行列交叉定位模式"""
        row_anchor_cfg = rule.get('row_anchor', {})
        col_anchor_cfg = rule.get('col_anchor', {})
        
        row_anchor = self._find_anchor(ws, {**row_anchor_cfg, 'search_in': row_anchor_cfg.get('search_in', 'first_col')})
        col_anchor = self._find_anchor(ws, {**col_anchor_cfg, 'search_in': col_anchor_cfg.get('search_in', 'first_row')})
        
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

# ---------------------------- v3.1 检查插件框架 ----------------------------
class CheckPlugin:
    """检查插件基类"""
    name = ""
    description = ""
    
    def __init__(self, config=None):
        self.config = config or {}
    
    def check(self, old_data, new_data, context=None):
        """ 执行检查 返回: [{'type': '...', 'desc': '...', 'severity': 'warning|error|info'}] """
        raise NotImplementedError

class MeanDeviationPlugin(CheckPlugin):
    """均值偏差检查插件：对比新旧均值，超过阈值则告警"""
    name = "mean_deviation"
    description = "检查均值偏差是否超过阈值"
    
    def check(self, old_data, new_data, context=None):
        results = []
        threshold = self.config.get('threshold', 0.05)  # 默认5%
        
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
    """参数锁定检查插件：检查关键参数是否被修改"""
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
    """范围检查插件：检查值是否在LSL/USL范围内"""
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

# 插件注册表
PLUGIN_REGISTRY = {
    'mean_deviation': MeanDeviationPlugin,
    'param_lock': ParamLockPlugin,
    'range_check': RangeCheckPlugin,
}

class PluginManager:
    """插件管理器：加载配置、执行插件"""
    def __init__(self, config_file=None):
        self.plugins = []
        self.locator = None
        if config_file and os.path.isfile(config_file):
            self.load_config(config_file)
    
    def load_config(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 加载定位规则
            if 'locator_rules' in data:
                self.locator = DataLocator()
                self.locator.rules = data['locator_rules']
            
            # 加载插件配置
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
        """执行所有插件检查，返回差异列表"""
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

# ---------------------------- 对比引擎 ----------------------------
class OpenpyxlComparer:
    def __init__(self, old_path, new_path, log_callback=None, progress_callback=None, check_options=None, plugin_manager=None):
        self.old_path = old_path
        self.new_path = new_path
        self.log = log_callback if log_callback else print
        self.progress = progress_callback if progress_callback else lambda v,s: None
        self.diffs = []
        self.sheet_diffs = []
        self.stats = {'total_cells':0, 'diff_cells':0, 'added_sheets':[], 'removed_sheets':[], 'images_diff':0}
        self.old_rich = {}
        self.new_rich = {}
        self.check_options = dict(DEFAULT_CHECK_OPTIONS)
        if check_options:
            self.check_options.update(check_options)
        self.plugin_manager = plugin_manager

    def run(self):
        opts = self.check_options
        enabled = [CHECK_OPTION_LABELS[k] for k, v in opts.items() if v]
        disabled = [CHECK_OPTION_LABELS[k] for k, v in opts.items() if not v]
        if disabled:
            self.log(f"已关闭 {len(disabled)} 项检测: {', '.join(disabled)}")

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
            self.progress(25 + int(55 * idx / total), f"对比 {sheet_name}...")
            self.log(f"正在对比 {sheet_name} ({idx}/{total})")
            old_ws = old_wb[sheet_name]
            new_ws = new_wb[sheet_name]
            self._compare_worksheet(old_ws, new_ws, sheet_name)

        # v3.1: 执行插件检查
        if self.plugin_manager and self.plugin_manager.plugins:
            self.progress(85, "执行数据检查插件...")
            self.log(f"执行 {len(self.plugin_manager.plugins)} 个检查插件...")
            plugin_results = self.plugin_manager.run_checks(old_wb, new_wb, self.log)
            for diff in plugin_results:
                self.diffs.append({
                    'sheet': '🔍 数据检查',
                    'address': diff.get('rule_name', ''),
                    'type': diff['type'],
                    'desc': diff['desc']
                })
            self.log(f"插件检查完成：发现 {len(plugin_results)} 个问题")

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
        max_row = max(old_ws.max_row or 1, new_ws.max_row or 1)
        max_col = max(old_ws.max_column or 1, new_ws.max_column or 1)
        self.stats['total_cells'] += max_row * max_col

        old_sheet_rich = self.old_rich.get(sheet_name, {})
        new_sheet_rich = self.new_rich.get(sheet_name, {})
        opts = self.check_options
        any_fmt = any(opts.get(k, True) for k in ['font', 'fill', 'border', 'alignment', 'number_format'])

        # v3.1 提速：使用 iter_rows 批量读取
        old_iter = old_ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col)
        new_iter = new_ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col)

        for row_idx, (old_row, new_row) in enumerate(zip_longest(old_iter, new_iter, fillvalue=()), start=1):
            for col_idx, (old_cell, new_cell) in enumerate(zip_longest(
                    old_row if old_row else (), 
                    new_row if new_row else (), 
                    fillvalue=None), start=1):
                
                # 处理一边没有cell的情况
                if old_cell is None or new_cell is None:
                    # 创建虚拟空cell
                    from openpyxl.cell.cell import Cell
                    if old_cell is None:
                        old_cell = Cell(old_ws, row=row_idx, column=col_idx)
                    if new_cell is None:
                        new_cell = Cell(new_ws, row=row_idx, column=col_idx)
                
                ref = cell_address(col_idx, row_idx)

                # 提速：快速跳过两边都完全为空的单元格
                if old_cell.value is None and new_cell.value is None:
                    if not any_fmt and ref not in old_sheet_rich and ref not in new_sheet_rich:
                        continue

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

        if opts.get('merged_cells', True):
            self._compare_merged_cells(old_ws, new_ws, sheet_name)
        if opts.get('row_height', True) or opts.get('col_width', True):
            self._compare_row_col_dimensions(old_ws, new_ws, sheet_name)
        if opts.get('images', True):
            self._compare_images(old_ws, new_ws, sheet_name)
        if opts.get('conditional_format', True):
            self._compare_conditional_formats(old_ws, new_ws, sheet_name)

    def _get_cell_diff(self, old_cell, new_cell, ref, old_runs, new_runs):
        opts = self.check_options

        if opts.get('rich_text', True) and (old_runs is not None or new_runs is not None):
            r1 = old_runs if old_runs is not None else (
                [(str(old_cell.value), None)] if old_cell.value is not None else [])
            r2 = new_runs if new_runs is not None else (
                [(str(new_cell.value), None)] if new_cell.value is not None else [])
            rich_diff = compare_rich_text_runs(r1, r2)
            if rich_diff:
                return {'type': '内容(含富文本)', 'desc': rich_diff}

        old_val = old_cell.value
        new_val = new_cell.value

        old_is_formula = isinstance(old_val, str) and old_val.startswith('=')
        new_is_formula = isinstance(new_val, str) and new_val.startswith('=')

        if old_is_formula and new_is_formula and old_val == new_val:
            pass
        elif old_val != new_val:
            old_str = f"{old_val}" if old_val is not None else "(空)"
            new_str = f"{new_val}" if new_val is not None else "(空)"
            if old_is_formula or new_is_formula:
                if opts.get('formula', True):
                    return {'type': '公式变化', 'desc': f'{old_str} → {new_str}'}
            else:
                if opts.get('value', True):
                    return {'type': '值变化', 'desc': f'{old_str} → {new_str}'}

        if not any(opts.get(k, True) for k in ['font', 'fill', 'border', 'alignment', 'number_format']):
            return None

        descs = []
        if opts.get('font', True):
            font_diff = self._cmp_font(old_cell.font, new_cell.font)
            if font_diff:
                descs.append(f"字体: {font_diff}")
        if opts.get('fill', True):
            fill_diff = self._cmp_fill(old_cell.fill, new_cell.fill)
            if fill_diff:
                descs.append(f"填充: {fill_diff}")
        if opts.get('border', True):
            border_diff = self._cmp_border(old_cell.border, new_cell.border)
            if border_diff:
                descs.append(f"边框: {border_diff}")
        if opts.get('alignment', True):
            align_diff = self._cmp_alignment(old_cell.alignment, new_cell.alignment)
            if align_diff:
                descs.append(f"对齐: {align_diff}")
        if opts.get('number_format', True):
            old_fmt = old_cell.number_format if old_cell.number_format is not None else 'General'
            new_fmt = new_cell.number_format if new_cell.number_format is not None else 'General'
            if old_fmt != new_fmt:
                descs.append(f"数字格式: {old_fmt} → {new_fmt}")

        if descs:
            return {'type': '格式变化', 'desc': '; '.join(descs)}
        return None

    def _cmp_font(self, f1, f2):
        changes = []
        n1 = f1.name if f1.name is not None else ''
        n2 = f2.name if f2.name is not None else ''
        if n1 != n2:
            changes.append(f"字体名: {n1 or '默认'}→{n2 or '默认'}")
        s1 = f1.size if f1.size is not None else 0
        s2 = f2.size if f2.size is not None else 0
        if s1 != s2:
            changes.append(f"大小: {s1}→{s2}")
        b1 = f1.bold if f1.bold is not None else False
        b2 = f2.bold if f2.bold is not None else False
        if b1 != b2:
            changes.append(f"粗体: {b1}→{b2}")
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
            self.diffs.append({'sheet':sheet_name,'address':addr.split(':')[0],'type':'合并新增','desc':f'新增合并区域 {addr}'})
        for addr in old_merged - new_merged:
            self.diffs.append({'sheet':sheet_name,'address':addr.split(':')[0],'type':'合并删除','desc':f'删除合并区域 {addr}'})

    def _compare_row_col_dimensions(self, old_ws, new_ws, sheet_name):
        opts = self.check_options
        if opts.get('row_height', True):
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
        if opts.get('col_width', True):
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
                    self.log(f"图片提取异常: {e}")
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

# ---------------------------- GUI ----------------------------
class DiffViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异对比工具 v3.1")
        self.root.geometry("1100x750")

        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.topmost = tk.BooleanVar(value=False)
        self.check_options = dict(DEFAULT_CHECK_OPTIONS)
        self.plugin_manager = None  # v3.1
        self.config_file = None

        # ========== 顶部工具栏 ==========
        toolbar = tb.Frame(root, padding=5)
        toolbar.pack(fill='x')
        toolbar.columnconfigure(5, weight=1)

        self.start_btn = tb.Button(toolbar, text="开始对比", bootstyle=INFO, width=8, command=self.start_compare)
        self.start_btn.grid(row=0, column=0, rowspan=2, sticky='nsew', padx=2, pady=1)

        self.jump_btn = tb.Button(toolbar, text="跳转", bootstyle=SECONDARY, width=6, command=self.jump_to_selected, state='disabled')
        self.jump_btn.grid(row=0, column=1, rowspan=2, sticky='nsew', padx=2, pady=1)

        self.settings_btn = tb.Button(toolbar, text="⚙ 检测设置", bootstyle="outline", width=10, command=self.open_check_options)
        self.settings_btn.grid(row=0, column=2, rowspan=2, sticky='nsew', padx=2, pady=1)

        # v3.1: 加载配置按钮
        self.config_btn = tb.Button(toolbar, text="📍 加载规则", bootstyle="outline", width=10, command=self.load_plugin_config)
        self.config_btn.grid(row=0, column=3, rowspan=2, sticky='nsew', padx=2, pady=1)

        tb.Separator(toolbar, orient='vertical').grid(row=0, column=4, rowspan=2, sticky='ns', padx=8)

        path_frame = tb.Frame(toolbar)
        path_frame.grid(row=0, column=5, rowspan=2, sticky='nsew', padx=(0, 10))
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
            row=0, column=6, rowspan=2, padx=(0, 5), pady=2, sticky='w')

        self.progress = tb.Progressbar(root, mode='determinate', bootstyle=PRIMARY)
        self.progress.pack(fill='x', padx=5, pady=(0, 5))

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

        bottom_frame = tb.Frame(root, padding=5)
        bottom_frame.pack(fill='x')
        bottom_frame.columnconfigure(0, weight=0)
        bottom_frame.columnconfigure(1, weight=1)

        detailf = tb.Labelframe(bottom_frame, text="差异详情", padding=5, bootstyle=INFO)
        detailf.grid(row=0, column=0, sticky='nsew', padx=(0, 3))
        detailf.grid_propagate(True)
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
        p = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if p: var.set(p)

    def toggle_topmost(self):
        self.root.attributes('-topmost', self.topmost.get())

    def open_check_options(self):
        dlg = CheckOptionsDialog(self.root, self.check_options)
        if dlg.result is not None:
            self.check_options = dlg.result
            enabled_count = sum(1 for v in self.check_options.values() if v)
            total_count = len(self.check_options)
            self.settings_btn.configure(text=f"⚙ 检测设置 ({enabled_count}/{total_count})")
            self.log(f"检测设置已更新：{enabled_count}/{total_count} 项已启用")

    def load_plugin_config(self):
        """v3.1: 加载插件配置文件"""
        filepath = filedialog.askopenfilename(
            title="选择检查规则配置文件",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not filepath:
            return
        
        self.plugin_manager = PluginManager(filepath)
        if self.plugin_manager.load_config(filepath):
            self.config_file = filepath
            rule_count = len(self.plugin_manager.locator.rules) if self.plugin_manager.locator else 0
            plugin_count = len(self.plugin_manager.plugins)
            self.config_btn.configure(text=f"📍 规则已加载 ({plugin_count})")
            self.log(f"规则配置加载成功：{rule_count} 条定位规则，{plugin_count} 个检查插件")
        else:
            messagebox.showerror("错误", "规则配置文件加载失败，请检查文件格式")
            self.plugin_manager = None

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

        current_opts = dict(self.check_options)
        pm = self.plugin_manager

        def worker():
            try:
                comparer = OpenpyxlComparer(old, new, self.log, self.update_progress,
                                           check_options=current_opts, plugin_manager=pm)
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
        
        # v3.1: 插件检查结果单独分组
        plugin_diffs = [d for d in diffs if d['sheet'] == '🔍 数据检查']
        normal_diffs = [d for d in diffs if d['sheet'] != '🔍 数据检查']
        
        if plugin_diffs:
            pn = self.tree.insert('','end',text='🔍 数据检查结果',open=True)
            for d in plugin_diffs:
                node = self.tree.insert(pn,'end',text=d['desc'][:80],values=(d['address'],d['type']))
                self.diff_items.append((node, {'type':'cell','data':d}))
        
        dmap = {}
        for d in normal_diffs:
            dmap.setdefault(d['sheet'],[]).append(d)
        for sname, items in sorted(dmap.items()):
            pn = self.tree.insert('','end',text=f"📄 {sname}",open=True)
            for d in items:
                node = self.tree.insert(pn,'end',text=d['desc'][:80],values=(d['address'],d['type']))
                self.diff_items.append((node, {'type':'cell','data':d}))
        
        self.jump_btn.configure(state='normal')
        enabled = sum(1 for v in self.check_options.values() if v)
        total = len(self.check_options)
        plugin_info = f"，{len(plugin_diffs)} 个插件告警" if plugin_diffs else ""
        self.log(f"树形列表已加载，单击查看详情，双击跳转（启用 {enabled}/{total} 项检测{plugin_info}）")

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

        # 插件检查结果不支持跳转
        if sheet == '🔍 数据检查':
            self.log("数据检查结果不支持跳转")
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
