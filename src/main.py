"""
Excel 差异对比工具（完整版）
对比两个 Excel 所有内容、格式、公式、合并单元格、行高列宽、条件格式等，基于新文件保留差异并添加注释。
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
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

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
            'renamed_sheets': {}
        }

    def compare_and_save(self, output_path):
        """执行全部对比并保存结果"""
        self.progress(5, "加载文件中...")
        self.log("正在加载工作簿...")
        old_wb = load_workbook(self.old_path)
        new_wb = load_workbook(self.new_path)
        result_wb = load_workbook(self.new_path)

        # 1. Sheet 层面差异
        self._compare_sheets(old_wb, new_wb, result_wb)

        # 2. 处理共有Sheet
        common = set(old_wb.sheetnames) & set(new_wb.sheetnames)
        total = len(common)
        for idx, sheet_name in enumerate(common, 1):
            self.progress(10 + int(80 * idx / total), f"对比 {sheet_name} ...")
            self.log(f"正在对比Sheet: {sheet_name} ({idx}/{total})")
            old_ws = old_wb[sheet_name]
            new_ws = new_wb[sheet_name]
            result_ws = result_wb[sheet_name]
            self._compare_worksheet(old_ws, new_ws, result_ws)
            # 行高列宽对比
            self._compare_row_col_dimensions(old_ws, new_ws, result_ws)
            # 条件格式对比
            self._compare_conditional_formatting(old_ws, new_ws, result_ws)

        # 3. 生成汇总报告
        self.progress(95, "生成汇总...")
        self._add_summary(result_wb)

        # 4. 保存
        self.progress(98, "保存文件...")
        result_wb.save(output_path)
        self.progress(100, "完成")
        self.log(f"对比完成，结果已保存至: {output_path}")
        self.log(f"统计：共 {self.stats['total_cells']} 单元格，{self.stats['diff_cells']} 处差异")

    def _compare_sheets(self, old_wb, new_wb, result_wb):
        old_set = set(old_wb.sheetnames)
        new_set = set(new_wb.sheetnames)
        self.stats['added_sheets'] = sorted(new_set - old_set)
        self.stats['removed_sheets'] = sorted(old_set - new_set)
        if self.stats['added_sheets']:
            self.log(f"新增Sheet: {', '.join(self.stats['added_sheets'])}")
        if self.stats['removed_sheets']:
            self.log(f"删除Sheet: {', '.join(self.stats['removed_sheets'])}")
        for name in self.stats['added_sheets']:
            if name in result_wb.sheetnames:
                ws = result_wb[name]
                if ws.max_row >= 1 and ws.max_column >= 1:
                    cell = ws.cell(1, 1)
                    self._append_comment(cell, "[新增Sheet，旧版中不存在]")

    def _compare_worksheet(self, old_ws, new_ws, result_ws):
        """对比单个工作表的所有差异，并清空完全相同的单元格"""
        max_row = max(old_ws.max_row, new_ws.max_row)
        max_col = max(old_ws.max_column, new_ws.max_column)
        self.stats['total_cells'] += max_row * max_col

        diff_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                old_cell = old_ws.cell(row, col)
                new_cell = new_ws.cell(row, col)
                result_cell = result_ws.cell(row, col)

                # 跳过合并单元格中的非左上角单元格
                if isinstance(result_cell, MergedCell):
                    continue

                # 检查是否完全相同
                if self._is_identical(old_cell, new_cell, old_ws, new_ws):
                    # 完全相同 -> 清空结果单元格，恢复默认样式
                    result_cell.value = None
                    result_cell.font = Font()
                    result_cell.fill = PatternFill()
                    result_cell.border = Border()
                    result_cell.alignment = Alignment()
                    result_cell.number_format = 'General'
                    result_cell.comment = None
                else:
                    self.stats['diff_cells'] += 1
                    # 保留新单元格的值和格式，高亮并添加差异注释
                    result_cell.fill = diff_fill
                    diff_desc = self._describe_diffs(old_cell, new_cell, old_ws, new_ws)
                    self._append_comment(result_cell, diff_desc)

        # 合并单元格差异
        self._compare_merged_cells(old_ws, new_ws, result_ws)

    def _is_identical(self, old_cell, new_cell, old_ws, new_ws):
        if old_cell.value != new_cell.value:
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
        """生成详细的差异说明，支持多差异同时显示"""
        diffs = []
        # 值差异（包括公式）
        if old_cell.value != new_cell.value:
            old_val = old_cell.value if old_cell.value is not None else "(空)"
            new_val = new_cell.value if new_cell.value is not None else "(空)"
            if (isinstance(old_cell.value, str) and old_cell.value.startswith('=')) or \
               (isinstance(new_cell.value, str) and new_cell.value.startswith('=')):
                diffs.append(f"公式: {old_val} → {new_val}")
            else:
                diffs.append(f"值: {old_val} → {new_val}")

        # 字体
        if not self._font_equal(old_cell.font, new_cell.font):
            f_diffs = []
            if old_cell.font.name != new_cell.font.name:
                f_diffs.append(f"字体名称: {old_cell.font.name}→{new_cell.font.name}")
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

        # 对齐方式
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

    def _compare_merged_cells(self, old_ws, new_ws, result_ws):
        old_merged = set(old_ws.merged_cells.ranges)
        new_merged = set(new_ws.merged_cells.ranges)
        for rng in new_merged - old_merged:
            cell = result_ws.cell(rng.min_row, rng.min_col)
            cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            self._append_comment(cell, "[新增合并区域]")

    def _compare_row_col_dimensions(self, old_ws, new_ws, result_ws):
        """对比行高和列宽，在有变化的行/列首单元格添加注释"""
        # 行高对比
        for row_idx in old_ws.row_dimensions:
            old_height = old_ws.row_dimensions[row_idx].height
            new_height = new_ws.row_dimensions[row_idx].height if row_idx in new_ws.row_dimensions else None
            if old_height != new_height:
                # 在A列对应行添加注释
                cell = result_ws.cell(row=row_idx, column=1)
                msg = f"行高变化: {old_height} → {new_height}"
                self._append_comment(cell, msg)
        for row_idx in new_ws.row_dimensions:
            if row_idx not in old_ws.row_dimensions:
                new_height = new_ws.row_dimensions[row_idx].height
                if new_height is not None:
                    cell = result_ws.cell(row=row_idx, column=1)
                    msg = f"行高新设置: {new_height}"
                    self._append_comment(cell, msg)

        # 列宽对比
        for col_letter in old_ws.column_dimensions:
            old_width = old_ws.column_dimensions[col_letter].width
            new_width = new_ws.column_dimensions[col_letter].width if col_letter in new_ws.column_dimensions else None
            if old_width != new_width:
                cell = result_ws.cell(row=1, column=col_letter_idx(col_letter))
                msg = f"列宽变化: {old_width} → {new_width}"
                self._append_comment(cell, msg)
        for col_letter in new_ws.column_dimensions:
            if col_letter not in old_ws.column_dimensions:
                new_width = new_ws.column_dimensions[col_letter].width
                if new_width is not None:
                    cell = result_ws.cell(row=1, column=col_letter_idx(col_letter))
                    msg = f"列宽新设置: {new_width}"
                    self._append_comment(cell, msg)

    def _compare_conditional_formatting(self, old_ws, new_ws, result_ws):
        """对比条件格式，如果有变化则在A1单元格添加提示"""
        # 将条件格式规则转为字符串比较（简化处理）
        old_rules = [str(cf) for cf in old_ws.conditional_formatting]
        new_rules = [str(cf) for cf in new_ws.conditional_formatting]
        if old_rules != new_rules:
            cell = result_ws.cell(row=1, column=1)
            self._append_comment(cell, "条件格式有变更，请检查条件格式规则")

    def _append_comment(self, cell, text):
        """为单元格添加或追加注释，并设置较大的文本框尺寸"""
        if cell.comment:
            cell.comment.text += "\n" + text
        else:
            comment = Comment(text, "ExcelDiff")
            comment.width = 400  # 设置注释框宽度
            comment.height = 200  # 设置注释框高度
            cell.comment = comment

    # -------------------- 格式比较辅助方法 --------------------
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
        stats_rows = [
            ("检查单元格总数", self.stats['total_cells']),
            ("差异单元格数", self.stats['diff_cells']),
            ("新增Sheet", len(self.stats['added_sheets'])),
            ("删除Sheet", len(self.stats['removed_sheets'])),
        ]
        for i, (label, val) in enumerate(stats_rows, 8):
            ws.cell(row=i, column=1, value=label)
            ws.cell(row=i, column=2, value=val)
        if self.stats['added_sheets']:
            ws['A14'] = "新增Sheet列表:"
            ws['A14'].font = Font(color="006600", bold=True)
            for j, name in enumerate(self.stats['added_sheets'], 15):
                ws.cell(row=j, column=1, value=name).font = Font(color="006600")
        if self.stats['removed_sheets']:
            next_row = 15 + len(self.stats['added_sheets'])
            ws.cell(row=next_row, column=1, value="删除Sheet列表:").font = Font(color="CC0000", bold=True)
            for j, name in enumerate(self.stats['removed_sheets'], next_row + 1):
                ws.cell(row=j, column=1, value=name).font = Font(color="CC0000")
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 30


def col_letter_idx(letter):
    """将列字母转换为列号（如 A->1, Z->26）"""
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


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
