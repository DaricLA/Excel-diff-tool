"""
Excel 差异对比工具（最终修复版）
- 移除 ConditionalFormattingList 导入，使用内置 clear 方法
- 其他功能保持不变
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.font import nametofont
import threading
import os
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText

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
            'sheets_with_diff': set()
        }

    def compare_and_save(self, output_path):
        self.progress(5, "加载文件中...")
        self.log("正在加载工作簿...")
        old_wb = load_workbook(self.old_path)
        new_wb = load_workbook(self.new_path)
        result_wb = load_workbook(self.new_path)

        self._compare_sheets(old_wb, new_wb, result_wb)

        common = set(old_wb.sheetnames) & set(new_wb.sheetnames)
        total = len(common)
        for idx, sheet_name in enumerate(common, 1):
            self.progress(10 + int(80 * idx / total), f"对比 {sheet_name} ...")
            self.log(f"正在对比 Sheet: {sheet_name} ({idx}/{total})")
            old_ws = old_wb[sheet_name]
            new_ws = new_wb[sheet_name]
            result_ws = result_wb[sheet_name]

            has_diff = self._compare_worksheet(old_ws, new_ws, result_ws)
            has_dim_diff = self._compare_row_col_dimensions(old_ws, new_ws, result_ws)
            has_cf_diff = self._compare_conditional_formatting(old_ws, new_ws, result_ws)

            if has_diff or has_dim_diff or has_cf_diff:
                self.stats['sheets_with_diff'].add(sheet_name)

        # 删除无差异工作表
        for name in list(result_wb.sheetnames):
            if name not in self.stats['sheets_with_diff'] and name != "差异汇总" and name != "已删除的Sheet":
                del result_wb[name]
                self.log(f"删除无差异工作表: {name}")

        # 强制注释可见
        for ws in result_wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.comment:
                        cell.comment.visible = True

        self.progress(95, "生成汇总...")
        self._add_summary(result_wb)

        self.progress(98, "保存文件...")
        result_wb.save(output_path)
        self.progress(100, "完成")
        self.log(f"对比完成，结果保存至: {output_path}")
        self.log(f"统计：共检查 {self.stats['total_cells']} 单元格，{self.stats['diff_cells']} 处差异")

    def _compare_sheets(self, old_wb, new_wb, result_wb):
        old_set = set(old_wb.sheetnames)
        new_set = set(new_wb.sheetnames)
        self.stats['added_sheets'] = sorted(new_set - old_set)
        self.stats['removed_sheets'] = sorted(old_set - new_set)

        for name in self.stats['added_sheets']:
            if name in result_wb.sheetnames:
                ws = result_wb[name]
                if ws.max_row >= 1 and ws.max_column >= 1:
                    self._add_comment(ws.cell(1, 1), "【新增 Sheet，旧版中不存在】")
                self.stats['sheets_with_diff'].add(name)

        if self.stats['removed_sheets']:
            placeholder = "已删除的Sheet"
            if placeholder in result_wb.sheetnames:
                del result_wb[placeholder]
            ws = result_wb.create_sheet(placeholder)
            ws['A1'] = "以下 Sheet 在新版中被删除："
            ws['A1'].font = Font(bold=True, color="FF0000")
            for i, name in enumerate(self.stats['removed_sheets'], 2):
                cell = ws.cell(row=i, column=1)
                cell.value = name
                cell.font = Font(color="FF0000")
                self._add_comment(cell, "【旧版中存在，新版已删除】")
            self.stats['sheets_with_diff'].add(placeholder)

    def _compare_worksheet(self, old_ws, new_ws, result_ws):
        max_row = max(old_ws.max_row, new_ws.max_row)
        max_col = max(old_ws.max_column, new_ws.max_column)
        self.stats['total_cells'] += max_row * max_col
        diff_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        has_diff = False

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                old_cell = old_ws.cell(row, col)
                new_cell = new_ws.cell(row, col)
                result_cell = result_ws.cell(row, col)

                if isinstance(result_cell, MergedCell):
                    continue

                if self._is_identical(old_cell, new_cell, old_ws, new_ws):
                    result_cell.value = None
                    result_cell.font = Font()
                    result_cell.fill = PatternFill()
                    result_cell.border = Border()
                    result_cell.alignment = Alignment()
                    result_cell.number_format = 'General'
                    result_cell.comment = None
                else:
                    has_diff = True
                    self.stats['diff_cells'] += 1
                    result_cell.fill = diff_fill
                    diff_desc = self._describe_diffs(old_cell, new_cell, old_ws, new_ws)
                    self._add_comment(result_cell, diff_desc)

        has_merge_diff = self._compare_merged_cells(old_ws, new_ws, result_ws)
        return has_diff or has_merge_diff

    def _is_identical(self, old_cell, new_cell, old_ws, new_ws):
        if not self._value_equal(old_cell.value, new_cell.value):
            return False
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
        if self._cell_is_merged(old_ws, old_cell.row, old_cell.column) != \
           self._cell_is_merged(new_ws, new_cell.row, new_cell.column):
            return False
        return True

    def _describe_diffs(self, old_cell, new_cell, old_ws, new_ws):
        diffs = []
        if not self._value_equal(old_cell.value, new_cell.value):
            old_rich = isinstance(old_cell.value, CellRichText)
            new_rich = isinstance(new_cell.value, CellRichText)
            if old_rich or new_rich:
                diffs.append(self._rich_text_diff(old_cell.value, new_cell.value))
            elif (isinstance(old_cell.value, str) and old_cell.value.startswith('=')) or \
                 (isinstance(new_cell.value, str) and new_cell.value.startswith('=')):
                diffs.append(f"公式: {self._format_value(old_cell.value)} → {self._format_value(new_cell.value)}")
            else:
                diffs.append(f"值: {self._format_value(old_cell.value)} → {self._format_value(new_cell.value)}")

        if not self._font_equal(old_cell.font, new_cell.font):
            f = []
            if old_cell.font.name != new_cell.font.name:
                f.append(f"字体: {old_cell.font.name}→{new_cell.font.name}")
            if old_cell.font.size != new_cell.font.size:
                f.append(f"大小: {old_cell.font.size}→{new_cell.font.size}")
            if old_cell.font.bold != new_cell.font.bold:
                f.append(f"粗体: {old_cell.font.bold}→{new_cell.font.bold}")
            if old_cell.font.italic != new_cell.font.italic:
                f.append(f"斜体: {old_cell.font.italic}→{new_cell.font.italic}")
            if old_cell.font.underline != new_cell.font.underline:
                f.append(f"下划线: {old_cell.font.underline}→{new_cell.font.underline}")
            if self._get_font_color(old_cell.font) != self._get_font_color(new_cell.font):
                f.append("文字颜色变更")
            if f:
                diffs.append("字体: " + "; ".join(f))

        if not self._fill_equal(old_cell.fill, new_cell.fill):
            diffs.append("填充色/背景变更")

        if not self._border_equal(old_cell.border, new_cell.border):
            detail = []
            for side in ['left', 'right', 'top', 'bottom']:
                old_side = getattr(old_cell.border, side)
                new_side = getattr(new_cell.border, side)
                if old_side.style != new_side.style or old_side.color != new_side.color:
                    desc = f"{side}: {old_side.style}({old_side.color})→{new_side.style}({new_side.color})"
                    detail.append(desc)
            if detail:
                diffs.append("边框: " + "; ".join(detail))

        if not self._alignment_equal(old_cell.alignment, new_cell.alignment):
            a = []
            if old_cell.alignment.horizontal != new_cell.alignment.horizontal:
                a.append(f"水平: {old_cell.alignment.horizontal}→{new_cell.alignment.horizontal}")
            if old_cell.alignment.vertical != new_cell.alignment.vertical:
                a.append(f"垂直: {old_cell.alignment.vertical}→{new_cell.alignment.vertical}")
            if old_cell.alignment.wrap_text != new_cell.alignment.wrap_text:
                a.append(f"自动换行: {old_cell.alignment.wrap_text}→{new_cell.alignment.wrap_text}")
            if a:
                diffs.append("对齐: " + "; ".join(a))

        if old_cell.number_format != new_cell.number_format:
            diffs.append(f"数字格式: {old_cell.number_format}→{new_cell.number_format}")

        old_m = self._cell_is_merged(old_ws, old_cell.row, old_cell.column)
        new_m = self._cell_is_merged(new_ws, new_cell.row, new_cell.column)
        if old_m != new_m:
            diffs.append("合并单元格状态变更")

        if not diffs:
            diffs.append("未知差异")
        return "【与旧版差异】\n" + "\n".join(diffs)

    def _rich_text_diff(self, old_val, new_val):
        if not isinstance(old_val, CellRichText):
            old_val = CellRichText(old_val if old_val is not None else "")
        if not isinstance(new_val, CellRichText):
            new_val = CellRichText(new_val if new_val is not None else "")
        old_plain = str(old_val)
        new_plain = str(new_val)
        if old_plain != new_plain:
            return f"内容(含局部格式): {old_plain} → {new_plain}"
        lines = []
        for i, (t1, t2) in enumerate(zip(old_val, new_val)):
            if t1.text != t2.text or not self._font_equal(t1.font, t2.font):
                segment = t1.text if t1.text else "(空)"
                changes = []
                if t1.font.name != t2.font.name:
                    changes.append(f"字体名: {t1.font.name}→{t2.font.name}")
                if t1.font.size != t2.font.size:
                    changes.append(f"大小: {t1.font.size}→{t2.font.size}")
                if t1.font.bold != t2.font.bold:
                    changes.append(f"粗体: {t1.font.bold}→{t2.font.bold}")
                if t1.font.italic != t2.font.italic:
                    changes.append(f"斜体: {t1.font.italic}→{t2.font.italic}")
                if t1.font.underline != t2.font.underline:
                    changes.append(f"下划线: {t1.font.underline}→{t2.font.underline}")
                if self._get_font_color(t1.font) != self._get_font_color(t2.font):
                    changes.append("颜色变更")
                if changes:
                    lines.append(f"段{i+1} “{segment}”: {'; '.join(changes)}")
                else:
                    lines.append(f"段{i+1} “{segment}”: 文本变为 “{t2.text}”")
        if not lines:
            lines.append("局部格式有细微变化")
        return "富文本局部格式变更:\n" + "\n".join(lines)

    @staticmethod
    def _value_equal(v1, v2):
        if type(v1) != type(v2):
            return False
        if isinstance(v1, CellRichText) and isinstance(v2, CellRichText):
            if len(v1) != len(v2):
                return False
            for t1, t2 in zip(v1, v2):
                if t1.text != t2.text:
                    return False
                if not ExcelDiffEngine._font_equal(t1.font, t2.font):
                    return False
            return True
        return v1 == v2

    @staticmethod
    def _format_value(val):
        if val is None:
            return "(空)"
        if isinstance(val, CellRichText):
            return str(val)
        return str(val)

    @staticmethod
    def _get_font_color(font):
        if font.color is None:
            return None
        try:
            return font.color.rgb
        except AttributeError:
            return str(font.color)

    @staticmethod
    def _fill_equal(f1, f2):
        if f1.fill_type != f2.fill_type:
            return False
        def get_rgb(fill, attr):
            color = getattr(fill, attr)
            if color is None:
                return None
            try:
                return color.rgb
            except AttributeError:
                return str(color)
        return (get_rgb(f1, 'start_color') == get_rgb(f2, 'start_color') and
                get_rgb(f1, 'end_color') == get_rgb(f2, 'end_color'))

    @staticmethod
    def _font_equal(f1, f2):
        return (f1.name == f2.name and f1.size == f2.size and f1.bold == f2.bold and
                f1.italic == f2.italic and f1.underline == f2.underline and
                ExcelDiffEngine._get_font_color(f1) == ExcelDiffEngine._get_font_color(f2))

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
        for rng in ws.merged_cells.ranges:
            if (rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col):
                return True
        return False

    def _compare_merged_cells(self, old_ws, new_ws, result_ws):
        old = set(old_ws.merged_cells.ranges)
        new = set(new_ws.merged_cells.ranges)
        added = new - old
        for rng in added:
            cell = result_ws.cell(rng.min_row, rng.min_col)
            cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            self._add_comment(cell, "【新增合并区域】")
        return len(added) > 0 or len(old - new) > 0

    def _compare_row_col_dimensions(self, old_ws, new_ws, result_ws):
        changed = False
        for row_idx in old_ws.row_dimensions:
            oh = old_ws.row_dimensions[row_idx].height
            nh = new_ws.row_dimensions[row_idx].height if row_idx in new_ws.row_dimensions else None
            if oh != nh:
                changed = True
                self._add_comment(result_ws.cell(row=row_idx, column=1), f"行高变化: {oh} → {nh}")
        for row_idx in new_ws.row_dimensions:
            if row_idx not in old_ws.row_dimensions:
                nh = new_ws.row_dimensions[row_idx].height
                if nh is not None:
                    changed = True
                    self._add_comment(result_ws.cell(row=row_idx, column=1), f"行高新设置: {nh}")
        for col_letter in old_ws.column_dimensions:
            ow = old_ws.column_dimensions[col_letter].width
            nw = new_ws.column_dimensions[col_letter].width if col_letter in new_ws.column_dimensions else None
            if ow != nw:
                changed = True
                col = column_index_from_string(col_letter)
                self._add_comment(result_ws.cell(row=1, column=col), f"列宽变化: {ow} → {nw}")
        for col_letter in new_ws.column_dimensions:
            if col_letter not in old_ws.column_dimensions:
                nw = new_ws.column_dimensions[col_letter].width
                if nw is not None:
                    changed = True
                    col = column_index_from_string(col_letter)
                    self._add_comment(result_ws.cell(row=1, column=col), f"列宽新设置: {nw}")
        return changed

    def _compare_conditional_formatting(self, old_ws, new_ws, result_ws):
        old_cfs = list(old_ws.conditional_formatting)
        new_cfs = list(new_ws.conditional_formatting)

        def serialize(cf):
            rules = []
            for r in cf.rules:
                rules.append((r.type, r.priority, str(r.dxf), str(r.formula)))
            return (str(cf.sqref), tuple(rules))

        old_set = {serialize(c) for c in old_cfs}
        new_set = {serialize(c) for c in new_cfs}

        if old_set == new_set:
            # 完全相同：清空条件格式（使用内置 clear 方法）
            result_ws.conditional_formatting.clear()
            return False
        else:
            for cf in new_cfs:
                ranges = cf.sqref
                for rng_str in ranges.split():
                    if ':' in rng_str:
                        start = rng_str.split(':')[0]
                    else:
                        start = rng_str
                    from openpyxl.utils import coordinate_to_tuple
                    try:
                        row, col = coordinate_to_tuple(start)
                        cell = result_ws.cell(row=row, column=col)
                        self._add_comment(cell, "【条件格式有变更】")
                    except:
                        pass
            return True

    def _add_comment(self, cell, text):
        if cell.comment:
            cell.comment.text += "\n" + text
            lines = cell.comment.text.count("\n") + 1
            cell.comment.width = 350
            cell.comment.height = max(120, lines * 18)
        else:
            comment = Comment(text, "ExcelDiff")
            comment.visible = True
            lines = text.count("\n") + 1
            comment.width = 350
            comment.height = max(120, lines * 18)
            cell.comment = comment

    def _add_summary(self, result_wb):
        if "差异汇总" in result_wb.sheetnames:
            del result_wb["差异汇总"]
        ws = result_wb.create_sheet("差异汇总", 0)
        ws['A1'] = "Excel差异对比汇总"
        ws['A1'].font = Font(size=16, bold=True)
        ws.merge_cells('A1:C1')

        ws['A3'] = "对比时间:"
        ws['B3'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws['A4'] = "旧版文件:"
        ws['B4'] = os.path.basename(self.old_path)
        ws['A5'] = "新版文件:"
        ws['B5'] = os.path.basename(self.new_path)

        ws['A7'] = "统计项"
        ws['B7'] = "数量"
        ws['A7'].font = Font(bold=True)
        ws['B7'].font = Font(bold=True)

        green = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        def color(val):
            return green if val == 0 else yellow

        stats = [
            ("检查单元格总数", self.stats['total_cells']),
            ("差异单元格数", self.stats['diff_cells']),
            ("新增 Sheet", len(self.stats['added_sheets'])),
            ("删除 Sheet", len(self.stats['removed_sheets'])),
        ]
        for i, (label, val) in enumerate(stats, 8):
            ws.cell(row=i, column=1, value=label)
            b = ws.cell(row=i, column=2, value=val)
            b.fill = color(val)

        row = 15
        if self.stats['added_sheets']:
            ws.cell(row=row, column=1, value="新增 Sheet 列表:").font = Font(color="006600", bold=True)
            for name in self.stats['added_sheets']:
                row += 1
                ws.cell(row=row, column=1, value=name).font = Font(color="006600")
        if self.stats['removed_sheets']:
            row += 2
            ws.cell(row=row, column=1, value="删除 Sheet 列表:").font = Font(color="CC0000", bold=True)
            for name in self.stats['removed_sheets']:
                row += 1
                ws.cell(row=row, column=1, value=name).font = Font(color="CC0000")

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 18


# ---------------------------- GUI ----------------------------
class ExcelDiffApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异对比工具")
        self.root.geometry("700x450")

        default_font = ("微软雅黑", 10)
        self.root.option_add("*Font", default_font)

        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill='both', expand=True)

        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky='ns', padx=(0, 10))
        left_frame.grid_rowconfigure(0, weight=1)
        left_frame.grid_rowconfigure(1, weight=1)

        self.compare_btn = tk.Button(
            left_frame,
            text="开始\n对比",
            font=("微软雅黑", 14, "bold"),
            bg="#0078D7", fg="white",
            width=10, height=4,
            relief='raised',
            command=self.start
        )
        self.compare_btn.grid(row=1, column=0, sticky='n')

        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky='nsew')
        right_frame.columnconfigure(1, weight=1)

        ttk.Label(right_frame, text="旧版文件:", font=("微软雅黑", 10)).grid(row=0, column=0, sticky='w', pady=5)
        self.old_path = tk.StringVar()
        old_entry = ttk.Entry(right_frame, textvariable=self.old_path, font=("微软雅黑", 10))
        old_entry.grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Button(right_frame, text="浏览", command=lambda: self.browse(self.old_path)).grid(row=0, column=2)

        ttk.Label(right_frame, text="新版文件:", font=("微软雅黑", 10)).grid(row=1, column=0, sticky='w', pady=5)
        self.new_path = tk.StringVar()
        new_entry = ttk.Entry(right_frame, textvariable=self.new_path, font=("微软雅黑", 10))
        new_entry.grid(row=1, column=1, padx=5, sticky='ew')
        ttk.Button(right_frame, text="浏览", command=lambda: self.browse(self.new_path)).grid(row=1, column=2)

        self.progress = ttk.Progressbar(root, mode='determinate')
        self.progress.pack(fill='x', padx=10, pady=(5, 0))

        self.log_area = scrolledtext.ScrolledText(root, height=12, font=("微软雅黑", 10))
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
            self.root.after(0, lambda: self.on_complete(output, engine.stats['diff_cells']))
        except Exception as e:
            self.log(f"错误: {e}")
            messagebox.showerror("失败", str(e))
        finally:
            self.progress['value'] = 0

    def on_complete(self, output_path, diff_count):
        if messagebox.askyesno("完成", f"对比完成！\n差异单元格: {diff_count}\n是否打开结果文件？"):
            try:
                os.startfile(output_path)
            except Exception as e:
                messagebox.showerror("打开失败", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDiffApp(root)
    root.mainloop()
