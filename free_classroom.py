"""
A GUI application to find free classrooms based on an Excel timetable.
Supports both .xls and .xlsx files using pandas and Tkinter.
"""
from __future__ import annotations

import os
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Set

import pandas as pd
import sys
import tkinter as tk
from tkinter import messagebox

@dataclass
class CourseEntry:
    """Represents a single course scheduling record."""

    class_name: str
    student_count: Optional[int]
    course_id: str
    course_name: str
    teacher: str
    teacher_department: str
    weekday: int
    start_section: int
    end_section: int
    location: str
    weeks: str
    week_type: str
    note: str

    @classmethod
    def from_row(cls, row: Dict) -> "CourseEntry":
        """Create a CourseEntry from a pandas row dictionary."""
        weekday, start, end = parse_time_code(str(row.get("开课时间", "")).strip())
        return cls(
            class_name=str(row.get("上课班级", "")).strip(),
            student_count=cls._safe_int(row.get("选课人数")),
            course_id=str(row.get("开课编号", "")).strip(),
            course_name=str(row.get("开课课程", "")).strip(),
            teacher=str(row.get("授课教师", "")).strip(),
            teacher_department=str(row.get("教师所在院系", "")).strip(),
            weekday=weekday,
            start_section=start,
            end_section=end,
            location=str(row.get("上课地点", "")).strip(),
            weeks=str(row.get("上课周次", "")).strip(),
            week_type=str(row.get("单双周", "")).strip(),
            note=str(row.get("备注", "")).strip(),
        )

    @staticmethod
    def _safe_int(value) -> Optional[int]:
        try:
            return int(value)
        except (TypeError, ValueError):
            return None


def parse_time_code(code: str) -> tuple[int, int, int]:
    """Parse the five-digit time code into weekday and section range.
    The format is ABCDE where:
    - A: weekday (1-7)
    - BC: start section (two digits)
    - DE: end section (two digits)
    """

    if len(code) != 5 or not code.isdigit():
        raise ValueError(f"开课时间格式错误: {code}")

    weekday = int(code[0])
    start_section = int(code[1:3])
    end_section = int(code[3:])

    if weekday < 1 or weekday > 7:
        raise ValueError(f"星期数超出范围 (1-7): {weekday}")
    if start_section <= 0 or end_section <= 0 or start_section > end_section:
        raise ValueError(f"节次范围错误: {start_section}-{end_section}")

    return weekday, start_section, end_section


class ClassroomFinder:
    """Handle loading timetable data and querying free classrooms."""

    REQUIRED_COLUMNS = [
        "上课班级",
        "选课人数",
        "开课编号",
        "开课课程",
        "授课教师",
        "教师所在院系",
        "开课时间",
        "上课地点",
        "上课周次",
        "单双周",
        "备注",
    ]

    def __init__(self) -> None:
        self.entries: List[CourseEntry] = []
        self.classrooms: Set[str] = set()

    def load_excel(self, file_path: str) -> None:
        """Load course data from an Excel file."""
        if not file_path:
            raise FileNotFoundError("请先选择 Excel 课表文件。")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"未找到文件: {file_path}")

        df = pd.read_excel(file_path, header=2, dtype=str)
        missing = [col for col in self.REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            raise ValueError(f"缺少必要列: {', '.join(missing)}")

        entries: List[CourseEntry] = []
        classrooms: Set[str] = set()

        for _, row in df.iterrows():
            try:
                entry = CourseEntry.from_row(row)
            except ValueError as exc:
                # Skip invalid rows but continue processing others
                print(f"跳过无效行: {exc}")
                continue
            if entry.location:
                classrooms.add(entry.location)
            entries.append(entry)

        self.entries = entries
        self.classrooms = classrooms
    def parse_week_range(self, week_str: str) -> Set[int]:
        """Parse week ranges like '2-4, 8-10, 12-20' into a set of week numbers."""
        weeks = set()
        if not week_str:
            return weeks
        parts = week_str.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = part.split('-')
                weeks.update(range(int(start), int(end) + 1))
            else:
                weeks.add(int(part))
        return weeks
    def match_week_type(self, current_week: int, week_type: str) -> bool:
        """Check if the current week matches the course week_type."""
        if week_type == "全周":
            return True
        if week_type == "单周":
            return current_week % 2 == 1
        if week_type == "双周":
            return current_week % 2 == 0
        return True  # fallback

    def find_free_classrooms(self, weekday: int, start_section: int, end_section: int,current_week: int) -> List[str]:
        """Return a sorted list of free classrooms for the given time slot."""
        if start_section > end_section:
            raise ValueError("开始节次不能大于结束节次")
        occupied: Set[str] = set()
        for entry in self.entries:

        # 1. 不是同一天 → 跳过
            if entry.weekday != weekday:
                continue

        # 2. 当前周是否在课程周次中？
            course_weeks = self.parse_week_range(entry.weeks)
            if current_week not in course_weeks:
                continue  # 本周不上课，跳过

        # 3. 单周 / 双周是否匹配？
            if not self.match_week_type(current_week, entry.week_type):
                continue  # 不符合单双周，上课周次，跳过

        # 4. 节次是否冲突？
            if self._is_overlapping(entry.start_section, entry.end_section, start_section, end_section):
                occupied.add(entry.location)

        free_rooms = sorted(room for room in self.classrooms if room not in occupied)
        return free_rooms

    @staticmethod
    def _is_overlapping(a_start: int, a_end: int, b_start: int, b_end: int) -> bool:
        """Check if two time ranges overlap."""
        return not (a_end < b_start or b_end < a_start)


class ClassroomFinderApp:
    """Tkinter GUI for classroom querying."""

    SECTION_STARTS = [1, 3, 5, 7, 9]
    SECTION_ENDS = [2, 4, 6, 8, 10, 12]

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("空教室查询")
        self.finder = ClassroomFinder()

        self._create_widgets()
    def export_today(self):
        if not self.finder.entries:
            try:
                self.load_data()
            except Exception:
                return

        try:
            weekday = int(self.weekday_var.get())
            current_week = int(self.current_week_var.get())
        except ValueError:
            messagebox.showerror("输入错误", "星期与周次必须是整数")
            return

    # 节次区间
        sections = [
            (1, 2, "1-2节"),
            (3, 4, "3-4节"),
            (5, 6, "5-6节"),
            (7, 8, "7-8节"),
            (9, 10, "9-10节"),
        ]

    # ---- 多列紧凑排版函数 ----
        def format_compact(items, per_line=4):
            lines = []
            for i in range(0, len(items), per_line):
                line = "  ".join(items[i:i + per_line])
                lines.append(line)
            return "\n".join(lines)

        output_lines = []

    # ============================================================
    # 🌞 新增部分：整个上午（1-4节），交集 = 同时空的教室
    # ============================================================
        morning_1 = set(self.finder.find_free_classrooms(weekday, 1, 2, current_week))
        morning_2 = set(self.finder.find_free_classrooms(weekday, 3, 4, current_week))
        morning_all = sorted(morning_1 & morning_2)

        output_lines.append("整个上午：")
        if morning_all:
            output_lines.append(format_compact(morning_all, per_line=4))
        else:
            output_lines.append("无空教室")
        output_lines.append("")

    # ============================================================
    # 🌇 新增部分：整个下午（5-8节），交集 = 同时空的教室
    # ============================================================
        afternoon_1 = set(self.finder.find_free_classrooms(weekday, 5, 6, current_week))
        afternoon_2 = set(self.finder.find_free_classrooms(weekday, 7, 8, current_week))
        afternoon_all = sorted(afternoon_1 & afternoon_2)

        output_lines.append("整个下午：")
        if afternoon_all:
            output_lines.append(format_compact(afternoon_all, per_line=4))
        else:
            output_lines.append("无空教室")
        output_lines.append("")

    # ============================================================
    # 原来的每个节次输出保持不变
    # ============================================================
        for start, end, label in sections:
            free_rooms = self.finder.find_free_classrooms(
                weekday, start, end, current_week
            )

            output_lines.append(f"{label}：")

            if free_rooms:
                output_lines.append(format_compact(free_rooms, per_line=4))
            else:
                output_lines.append("无空教室")

            output_lines.append("")

    # ---- 保存文件 ----
        save_path = filedialog.asksaveasfilename(
            title="导出空教室（手机紧凑版）",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")],
        )

        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write("\n".join(output_lines))
            messagebox.showinfo("导出成功", f"已生成适合手机阅读的紧凑版空教室：\n{save_path}")

    def _create_widgets(self) -> None:
        main_frame = ttk.Frame(self.root, padding="12 12 12 12")
        main_frame.grid(row=0, column=0, sticky="NSEW")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    # 文件选择
        ttk.Label(main_frame, text="Excel 文件路径:").grid(row=0, column=0, sticky="W")
        self.file_path_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.file_path_var, width=40).grid(row=0, column=1, columnspan=2, sticky="WE", padx=5)
        ttk.Button(main_frame, text="选择文件", command=self.select_file).grid(row=0, column=3, sticky="E")

    # 星期
        ttk.Label(main_frame, text="星期 (1-7):").grid(row=1, column=0, sticky="W", pady=5)
        self.weekday_var = tk.StringVar(value="1")
        ttk.Combobox(main_frame, textvariable=self.weekday_var,
                 values=[str(i) for i in range(1, 8)], state="readonly", width=10).grid(row=1, column=1, sticky="W")

    # 开始节次
        ttk.Label(main_frame, text="开始节次:").grid(row=1, column=2, sticky="E")
        self.start_var = tk.StringVar(value=str(self.SECTION_STARTS[0]))
        ttk.Combobox(main_frame, textvariable=self.start_var,
                 values=[str(s) for s in self.SECTION_STARTS],
                 state="readonly", width=10).grid(row=1, column=3, sticky="W")

    # 结束节次
        ttk.Label(main_frame, text="结束节次:").grid(row=2, column=0, sticky="W", pady=5)
        self.end_var = tk.StringVar(value=str(self.SECTION_ENDS[0]))
        ttk.Combobox(main_frame, textvariable=self.end_var,
                 values=[str(e) for e in self.SECTION_ENDS],
                 state="readonly", width=10).grid(row=2, column=1, sticky="W")

    # 当前周次（修复重点）
        ttk.Label(main_frame, text="当前周次:").grid(row=2, column=2, sticky="E")
        self.current_week_var = tk.StringVar(value="1")
        ttk.Entry(main_frame, textvariable=self.current_week_var, width=10).grid(row=2, column=3, sticky="W")

    # 查询按钮
        ttk.Button(main_frame, text="查询空教室", command=self.query).grid(row=3, column=3, sticky="E", pady=10)
    # 导出按钮
        ttk.Button(main_frame, text="导出今日空教室", command=self.export_today).grid(row=3, column=0, sticky="W", pady=10)

    # 空教室列表
        ttk.Label(main_frame, text="空教室列表:").grid(row=4, column=0, sticky="W")
        self.result_text = tk.Text(main_frame, width=60, height=20)
        self.result_text.grid(row=5, column=0, columnspan=4, sticky="NSEW", pady=(5, 0))

    # grid 扩展
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)


    def select_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Excel 课表文件",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.load_data()

    def load_data(self) -> None:
        try:
            self.finder.load_excel(self.file_path_var.get())
            messagebox.showinfo("成功", "课表数据加载完成！")
        except Exception as exc:  # noqa: BLE001 - GUI friendly catch-all
            messagebox.showerror("错误", str(exc))

    def query(self) -> None:
        if not self.finder.entries:
            try:
                self.load_data()
            except Exception:
                return

        try:
            weekday = int(self.weekday_var.get())
            start = int(self.start_var.get())
            end = int(self.end_var.get())
            current_week = int(self.current_week_var.get())  # ← 这里要用 current_week_var

            results = self.finder.find_free_classrooms(
            weekday, start, end, current_week
        )
        except ValueError as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        self.result_text.delete("1.0", tk.END)
        if not results:
            self.result_text.insert(tk.END, "没有找到空教室。\n")
            return

        self.result_text.insert(tk.END, "\n".join(results))



def main() -> None:
    root = tk.Tk()
    app = ClassroomFinderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()