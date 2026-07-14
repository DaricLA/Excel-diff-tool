"""
Excel 差异对比工具（完整版）
支持：富文本（同一单元格内部分字符格式不同）、默认显示注释、行高列宽、条件格式、
      增减sheet注释标记、汇总统计颜色标注等。
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from datetime import datetime
from copy import copy

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Alignment, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock

# ---------------------------- 核心对比引擎 ----------------------------
class ExcelDiffEngine:
    def __init__(self, old_path, new_path, log_func=None, progress_func=None):
        self.old_path = old_path
        self.new_path = new_path
        self.log = log_func or print
        self.progress = progress_func or (lambda v, s: None)
        self.stats = {
            'total_cells': 0,
            'diff_cells': 0,
            'added_sheets': [],
            'removed_sheets': [],
            'cond_format_diffs': []   # 记录有条件格式差异的sheet
        }

    def compare_and_save(self, output_path):
        self.progress(5, "加载文件中...")
        self.log("正在加载工作簿...")
        old_wb = load_workbook(self.old_path)
        new_wb = load_workbook(self.new_path)
        result_wb = load_workbook(self.new_path)

        # 1. Sheet 差异
        self._compare_sheets(old_wb, new_wb, result_wb)

        # 2. 共有 Sheet 详细对比
        common = set(old_wb.sheetnames) & set(new_wb.sheetnames)
        total = len(common)
        for idx, sheet_name in enumerate(common, 1):
            self.progress(10 + int(80 * idx / total), f"对比 {sheet_name} ...")
            self.log(f"正在对比 Sheet: {sheet_name} ({idx}/{total})")
            old_ws = old_wb[sheet_name]
            new_ws = new_wb[sheet_name]
            result_ws = result_wb[sheet_name]
            self._compare_worksheet(old_ws, new_ws, result_ws)
            self._compare_row_col_dimensions(old_ws, new_ws, result_ws)
            self._compare_conditional_formatting(old_ws, new_ws, result_ws, sheet_name)

        # 3. 汇总报告
        self.progress(95, "生成汇总...")
        self._add_summary(result_wb)

        self.progress(98, "保存文件...")
        result_wb.save(output_path)
        self.progress(100, "完成")
        self.log(f"对比完成，结果保存至: {output_path}")
        self.log(f"统计：共检查 {self.stats['total_cells']} 单元格，发现 {self.stats['diff_cells']} 处差异")

    # ---------- Sheet 差异 ----------
    def _compare_sheets(self, old_wb, new_wb, result_wb):
        old_set = set(old_wb.sheetnames)
        new_set = set(new_wb.sheetnames)
        self.stats['added_sheets'] = sorted(new_set - old_set)
        self.stats['removed_sheets'] = sorted(old_set - new_set)

        # 为新增 sheet 在第一个单元格添加注释
        for name in self.stats['added_sheets']:
            if name in result_wb.sheetnames:
                ws = result_wb[name]
                if ws.max_row >= 1 and ws.max_column >= 1:
                    cell = ws.cell(1, 1)
                    self._add_comment(cell, "【新增 Sheet，旧版中不存在】")

        # 为删除的 sheet 在汇总中标记（汇总生成时处理），同时也插入占位 sheet 便于查看
        if self.stats['removed_sheets']:
            # 如果已存在占位 sheet 则删除重建
            placeholder_name = "已删除的Sheet"
            if placeholder_name in result_wb.sheetnames:
                del result_wb[placeholder_name]
            ws = result_wb.create_sheet(placeholder_name)
            ws['A1'] = "以下 Sheet 在新版中被删除："
            ws['A1'].font = Font(bold=True, color="FF0000")
            for i, name in enumerate(self.stats['removed_sheets'], 2):
                cell = ws.cell(row=i, column=1)
                cell.value = name
                cell.font = Font(color="FF0000")
                self._add_comment(cell, f"【旧版中存在，新版已删除】")

            self.log(f"删除 Sheet: {', '.join(self.stats['removed_sheets'])}")

    # ---------- 工作表单元格对比 ----------
    def _compare_worksheet(self, old_ws, new_ws, result_ws):
        max_row = max(old_ws.max_row, new_ws.max_row)
        max_col = max(old_ws.max_column, new_ws.max_column)
        self.stats['total_cells'] += max_row * max_col

        diff_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                old_cell = old_ws.cell(row, col)
                new_cell = new_ws.cell(row, col)
                result_cell = result_ws.cell(row, col)

                if isinstance(result_cell, MergedCell):
                    continue

                # 检查是否完全相同（含富文本格式）
                if self._is_identical(old_cell, new_cell, old_ws, new_ws):
                    # 完全相同则清空
                    result_cell.value = None
                    result_cell.font = Font()
                    result_cell.fill = PatternFill()
                    result_cell.border = Border()
                    result_cell.alignment = Alignment()
                    result_cell.number_format = 'General'
                    result_cell.comment = None
                else:
                    self.stats['diff_cells'] += 1
                    result_cell.fill = diff_fill
                    diff_desc = self._describe_diffs(old_cell, new_cell, old_ws, new_ws)
                    self._add_comment(result_cell, diff_desc)

        self._compare_merged_cells(old_ws, new_ws, result_ws)

    def _is_identical(self, old_cell, new_cell, old_ws, new_ws):
        # 比较值（需处理富文本）
        if not self._value_equal(old_cell.value, new_cell.value):
            return False
        # 比较单元格级别格式
        if not self._font_equal(old_cell.font, new_cell.font):
            return False
        if not self._fill_equal(old_cell.fill, new_cell.fill):
            return False
        if not self._border_equal(old_cell.border, new_cell.border):
            return False
        if not self._alignment_equal(old_cell.alignment, new_cell.alignment):
            return False
        if old_cell.number_format != new_cell.number_format:
            return False
        # 合并状态
        if self._cell_is_merged(old_ws, old_cell.row, old_cell.column) != \
           self._cell_is_merged(new_ws, new_cell.row, new_cell.column):
            return False
        return True

    def _describe_diffs(self, old_cell, new_cell, old_ws, new_ws):
        diffs = []
        # 值差异（含富文本）
        if not self._value_equal(old_cell.value, new_cell.value):
            old_val = self._format_value(old_cell.value)
            new_val = self._format_value(new_cell.value)
            # 区分富文本和普通值
            old_is_rich = isinstance(old_cell.value, CellRichText)
            new_is_rich = isinstance(new_cell.value, CellRichText)
            if old_is_rich or new_is_rich:
                diffs.append(f"内容(含局部格式): {old_val} → {new_val}")
            elif (isinstance(old_cell.value, str) and old_cell.value.startswith('=')) or \
                 (isinstance(new_cell.value, str) and new_cell.value.startswith('=')):
                diffs.append(f"公式: {old_val} → {new_val}")
            else:
                diffs.append(f"值: {old_val} → {new_val}")

        # 字体差异
        if not self._font_equal(old_cell.font, new_cell.font):
            f_diffs = []
            if old_cell.font.name != new_cell.font.name:
                f_diffs.append(f"字体: {old_cell.font.name}→{new_cell.font.name}")
            if old_cell.font.size != new_cell.font.size:
                f_diffs.append(f"大小: {old_cell.font.size}→{new_cell.font.size}")
            if old_cell.font.bold != new_cell.font.bold:
                f_diffs.append(f"粗体: {old_cell.font.bold}→{new_cell.font.bold}")
            if old_cell.font.italic != new_cell.font.italic:
                f_diffs.append(f"斜体: {old_cell.font.italic}→{new_cell.font.italic}")
            if old_cell.font.underline != new_cell.font.underline:
                f_diffs.append(f"下划线: {old_cell.font.underline}→{new_cell.font.underline}")
            if old_cell.font.color != new_cell.font.color:
                f_diffs.append("文字颜色变更")
            if f_diffs:
                diffs.append("字体: " + "; ".join(f_diffs))

        # 填充色
        if not self._fill_equal(old_cell.fill, new_cell.fill):
            diffs.append("填充色/背景变更")

        # 边框
        if not self._border_equal(old_cell.border, new_cell.border):
            diffs.append("边框样式变更")

        # 对齐
        if not self._alignment_equal(old_cell.alignment, new_cell.alignment):
            a_diffs = []
            if old_cell.alignment.horizontal != new_cell.alignment.horizontal:
                a_diffs.append(f"水平: {old_cell.alignment.horizontal}→{new_cell.alignment.horizontal}")
            if old_cell.alignment.vertical != new_cell.alignment.vertical:
                a_diffs.append(f"垂直: {old_cell.alignment.vertical}→{new_cell.alignment.vertical}")
            if old_cell.alignment.wrap_text != new_cell.alignment.wrap_text:
                a_diffs.append(f"自动换行: {old_cell.alignment.wrap_text}→{new_cell.alignment.wrap_text}")
            if a_diffs:
                diffs.append("对齐: " + "; ".join(a_diffs))

        # 数字格式
        if old_cell.number_format != new_cell.number_format:
            diffs.append(f"数字格式: {old_cell.number_format}→{new_cell.number_format}")

        # 合并状态
        old_merged = self._cell_is_merged(old_ws, old_cell.row, old_cell.column)
        new_merged = self._cell_is_merged(new_ws, new_cell.row, new_cell.column)
        if old_merged != new_merged:
            diffs.append("合并单元格状态变更")

        if not diffs:
            diffs.append("未知差异")
        return "【与旧版差异】\n" + "\n".join(diffs)

    # ---------- 富文本处理 ----------
    @staticmethod
    def _value_equal(val1, val2):
        """判断两个单元格值是否相同（考虑富文本）"""
        if type(val1) != type(val2):
            return False
        if isinstance(val1, CellRichText) and isinstance(val2, CellRichText):
            # 比较富文本：长度、每段文本和格式
            if len(val1) != len(val2):
                return False
            for t1, t2 in zip(val1, val2):
                if t1.text != t2.text:
                    return False
                if not ExcelDiffEngine._font_equal(t1.font, t2.font):
                    return False
            return True
        else:
            return val1 == val2

    @staticmethod
    def _format_value(val):
        """将值转换为适合显示的字符串"""
        if val is None:
            return "(空)"
        if isinstance(val, CellRichText):
            return str(val)  # 纯文本显示
        return str(val)

    # ---------- 合并单元格 ----------
    def _compare_merged_cells(self, old_ws, new_ws, result_ws):
        old_merged = set(old_ws.merged_cells.ranges)
        new_merged = set(new_ws.merged_cells.ranges)
        for rng in new_merged - old_merged:
            cell = result_ws.cell(rng.min_row, rng.min_col)
            cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            self._add_comment(cell, "【新增合并区域】")

    # ---------- 行高列宽 ----------
    def _compare_row_col_dimensions(self, old_ws, new_ws, result_ws):
        # 行高
        for row_idx in old_ws.row_dimensions:
            old_height = old_ws.row_dimensions[row_idx].height
            new_height = new_ws.row_dimensions[row_idx].height if row_idx in new_ws.row_dimensions else None
            if old_height != new_height:
                cell = result_ws.cell(row=row_idx, column=1)
                self._add_comment(cell, f"行高变化: {old_height} → {new_height}")
        for row_idx in new_ws.row_dimensions:
            if row_idx not in old_ws.row_dimensions:
                new_height = new_ws.row_dimensions[row_idx].height
                if new_height is not None:
                    cell = result_ws.cell(row=row_idx, column=1)
                    self._add_comment(cell, f"行高新设置: {new_height}")

        # 列宽
        for col_letter in old_ws.column_dimensions:
            old_width = old_ws.column_dimensions[col_letter].width
            new_width = new_ws.column_dimensions[col_letter].width if col_letter in new_ws.column_dimensions else None
            if old_width != new_width:
                col_idx = column_index_from_string(col_letter)
                cell = result_ws.cell(row=1, column=col_idx)
                self._add_comment(cell, f"列宽变化: {old_width} → {new_width}")
        for col_letter in new_ws.column_dimensions:
            if col_letter not in old_ws.column_dimensions:
                new_width = new_ws.column_dimensions[col_letter].width
                if new_width is not None:
                    col_idx = column_index_from_string(col_letter)
                    cell = result_ws.cell(row=1, column=col_idx)
                    self._add_comment(cell, f"列宽新设置: {new_width}")

    # ---------- 条件格式 ----------
    def _compare_conditional_formatting(self, old_ws, new_ws, result_ws, sheet_name):
        # 使用条件格式范围对象进行比较
        old_cfs = list(old_ws.conditional_formatting)
        new_cfs = list(new_ws.conditional_formatting)
        # 由于条件格式对象难以直接相等比较，将其序列化为可比较的字符串
        def serialize_cf(cf):
            rules_str = []
            for rule in cf.rules:
                # 提取规则的主要属性
                rules_str.append(f"type={rule.type},priority={rule.priority},"
                                 f"dxf={rule.dxf},formula={rule.formula}")
            return (str(cf.sqref), tuple(rules_str))

        old_set = {serialize_cf(cf) for cf in old_cfs}
        new_set = {serialize_cf(cf) for cf in new_cfs}

        if old_set != new_set:
            self.stats.setdefault('cond_format_diffs', []).append(sheet_name)
            # 在 A1 单元格标记
            cell = result_ws.cell(row=1, column=1)
            self._add_comment(cell, "【条件格式有变更，请检查条件格式规则】")

    # ---------- 注释工具（默认显示、大尺寸） ----------
    def _add_comment(self, cell, text):
        """添加或追加注释，设置自动尺寸并默认显示"""
        if cell.comment:
            # 追加内容
            cell.comment.text += "\n" + text
            # 重新设置尺寸（可根据内容行数略调）
            lines = cell.comment.text.count("\n") + 1
            cell.comment.width = 350
            cell.comment.height = max(150, lines * 20)
        else:
            comment = Comment(text, "ExcelDiff")
            comment.visible = True  # 默认显示，而不是悬停才显示
            lines = text.count("\n") + 1
            comment.width = 350
            comment.height = max(150, lines * 20)
            cell.comment = comment

    # ---------- 格式比较静态方法 ----------
    @staticmethod
    def _font_equal(f1, f2):
        return (f1.name == f2.name and f1.size == f2.size and f1.bold == f2.bold and
                f1.italic == f2.italic and f1.underline == f2.underline and f1.color == f2.color)

    @staticmethod
    def _fill_equal(f1, f2):
        return (f1.fill_type == f2.fill_type and f1.start_color == f2.start_color and
                f1.end_color == f2.end_color)

    @staticmethod
    def _border_equal(b1, b2):
        for side in ['left', 'right', 'top', 'bottom']:
            s1 = getattr(b1, side)
            s2 = getattr(b2, side)
            if s1.style != s2.style or s1.color != s2.color:
                return False
        return True

    @staticmethod
    def _alignment_equal(a1, a2):
        return (a1.horizontal == a2.horizontal and a1.vertical == a2.vertical and
                a1.wrap_text == a2.wrap_text)

    @staticmethod
    def _cell_is_merged(ws, row, col):
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
                return True
        return False

    # ---------- 汇总报告 ----------
    def _add_summary(self, result_wb):
        if "差异汇总" in result_wb.sheetnames:
            del result_wb["差异汇总"]
        ws = result_wb.create_sheet("差异汇总", 0)

        # 标题
        ws['A1'] = "Excel差异对比汇总"
        ws['A1'].font = Font(size=16, bold=True)
        ws.merge_cells('A1:C1')

        ws['A3'] = "对比时间:"
        ws['B3'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws['A4'] = "旧版文件:"
        ws['B4'] = os.path.basename(self.old_path)
        ws['A5'] = "新版文件:"
        ws['B5'] = os.path.basename(self.new_path)

        # 统计表头
        ws['A7'] = "统计项"
        ws['B7'] = "数量"
        ws['A7'].font = Font(bold=True)
        ws['B7'].font = Font(bold=True)

        # 颜色规则：0 用绿色，>0 用黄色
        green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        def color_fill(val):
            return green_fill if val == 0 else yellow_fill

        stats_data = [
            ("检查单元格总数", self.stats['total_cells']),
            ("差异单元格数", self.stats['diff_cells']),
            ("新增 Sheet", len(self.stats['added_sheets'])),
            ("删除 Sheet", len(self.stats['removed_sheets'])),
            ("有条件格式差异的 Sheet", len(self.stats.get('cond_format_diffs', [])))
        ]

        for i, (label, val) in enumerate(stats_data, 8):
            cell_a = ws.cell(row=i, column=1)
            cell_a.value = label
            cell_b = ws.cell(row=i, column=2)
            cell_b.value = val
            cell_b.fill = color_fill(val)

        # 列表信息
        row = 15
        if self.stats['added_sheets']:
            ws.cell(row=row, column=1, value="新增 Sheet 列表:").font = Font(color="006600", bold=True)
            for name in self.stats['added_sheets']:
                row += 1
                c = ws.cell(row=row, column=1, value=name)
                c.font = Font(color="006600")
                self._add_comment(c, "【新增 Sheet】")

        if self.stats['removed_sheets']:
            row += 2
            ws.cell(row=row, column=1, value="删除 Sheet 列表:").font = Font(color="CC0000", bold=True)
            for name in self.stats['removed_sheets']:
                row += 1
                c = ws.cell(row=row, column=1, value=name)
                c.font = Font(color="CC0000")
                self._add_comment(c, "【已删除，旧版中存在】")

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 18


# ---------------------------- GUI ----------------------------
class ExcelDiffApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异对比工具")
        self.root.geometry("650x500")
        frame = ttk.Frame(root, padding=10)
        frame.pack(fill='x')
        ttk.Label(frame, text="旧版文件:").grid(row=0, column=0, sticky='w')
        self.old_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.old_path, width=55).grid(row=0, column=1, padx=5)
        ttk.Button(frame, text="浏览", command=lambda: self.browse(self.old_path)).grid(row=0, column=2)
        ttk.Label(frame, text="新版文件:").grid(row=1, column=0, sticky='w', pady=5)
        self.new_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.new_path, width=55).grid(row=1, column=1, padx=5)
        ttk.Button(frame, text="浏览", command=lambda: self.browse(self.new_path)).grid(row=1, column=2)
        ttk.Button(frame, text="开始对比", command=self.start).grid(row=2, column=1, pady=10)
        self.progress = ttk.Progressbar(root, mode='determinate')
        self.progress.pack(fill='x', padx=10)
        self.log_area = scrolledtext.ScrolledText(root, height=15)
        self.log_area.pack(fill='both', expand=True, padx=10, pady=10)

    def browse(self, var):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            var.set(filename)

    def log(self, msg):
        self.log_area.insert('end', f"[{datetime.now():%H:%M:%S}] {msg}\n")
        self.log_area.see('end')
        self.root.update_idletasks()

    def update_progress(self, value, status=""):
        self.progress['value'] = value
        if status:
            self.log(status)
        self.root.update_idletasks()

    def start(self):
        old = self.old_path.get()
        new = self.new_path.get()
        if not old or not new:
            messagebox.showerror("错误", "请选择两个Excel文件")
            return
        output = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not output:
            return
        threading.Thread(target=self.run, args=(old, new, output), daemon=True).start()

    def run(self, old, new, output):
        try:
            engine = ExcelDiffEngine(old, new, log_func=self.log, progress_func=self.update_progress)
            engine.compare_and_save(output)
            messagebox.showinfo("完成", f"对比完成！\n差异单元格: {engine.stats['diff_cells']}\n结果保存至:\n{output}")
        except Exception as e:
            self.log(f"错误: {e}")
            messagebox.showerror("失败", str(e))
        finally:
            self.progress['value'] = 0


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDiffApp(root)
    root.mainloop()
