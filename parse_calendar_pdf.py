"""行事予定PDF（1年・2年・2国・3年列の〇/○・△・✕等）から授業日を集計するパーサ。"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from typing import BinaryIO

import pdfplumber

WEEKDAYS_JA = "月火水木金土日"
GRADES = ("1年", "2年", "2国", "3年")
# 授業日として数える記号（PDFでは主に 〇 U+3007、一部 ○ U+25CB）
CLASS_MARKS = frozenset({"〇", "○", "◯"})


@dataclass
class CalendarRow:
    year: int
    month: int
    day: int
    weekday: str
    marks: tuple[str, str, str, str]


def _parse_page_title(title_cell: str | None) -> tuple[int, int] | None:
    if not title_cell or not str(title_cell).strip():
        return None
    cell = str(title_cell).strip()

    m = re.search(r"2027年\s*(\d+)\s*月", cell)
    if m:
        return 2027, int(m.group(1))

    m = re.search(r"年度\s+2026年\s+(\d+)\s*月", cell)
    if m:
        return 2026, int(m.group(1))

    first_line = cell.split("\n")[0].strip()
    m = re.match(r"^(\d+)\s*月$", first_line)
    if m:
        month = int(m.group(1))
        if month >= 5:
            return 2026, month
        if month == 1:
            return 2027, 1
        if month in (2, 3):
            return 2027, month
        if month == 4:
            return 2026, 4

    return None


def _is_data_row(c0: str | None, c1: str | None) -> bool:
    if c0 is None or c1 is None:
        return False
    d = str(c0).strip()
    w = str(c1).strip()
    if not d.isdigit():
        return False
    day = int(d)
    if day < 1 or day > 31:
        return False
    return w in WEEKDAYS_JA


def _norm_mark(cell: str | None) -> str:
    if cell is None:
        return ""
    return str(cell).strip()


def parse_pdf(file: BinaryIO | str) -> tuple[list[CalendarRow], list[str]]:
    rows: list[CalendarRow] = []
    warnings: list[str] = []

    with pdfplumber.open(file) as pdf:
        for page_index, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if not tables:
                warnings.append(f"p.{page_index + 1}: 表が見つかりません")
                continue
            table = tables[0]
            if len(table) < 3:
                continue

            title_cell = table[0][0] if table[0] else None
            my = _parse_page_title(title_cell)
            if my is None:
                warnings.append(f"p.{page_index + 1}: 月を解釈できません: {title_cell!r}")
                continue
            year, month = my

            for r in table[2:]:
                if not r or len(r) < 6:
                    continue
                c0, c1 = r[0], r[1]
                if str(c0 or "").strip() == "備考":
                    break
                if not _is_data_row(c0, c1):
                    continue
                day = int(str(c0).strip())
                weekday = str(c1).strip()
                marks = (
                    _norm_mark(r[2]),
                    _norm_mark(r[3]),
                    _norm_mark(r[4]),
                    _norm_mark(r[5]),
                )
                if not all(marks):
                    continue
                try:
                    expected = WEEKDAYS_JA[date(year, month, day).weekday()]
                except ValueError:
                    warnings.append(f"p.{page_index + 1}: 無効な日付 {year}-{month}-{day}")
                    continue
                if expected != weekday:
                    warnings.append(
                        f"{year}-{month}-{day}: 曜日が暦と不一致 (表:{weekday} 暦:{expected})"
                    )

                rows.append(
                    CalendarRow(
                        year=year,
                        month=month,
                        day=day,
                        weekday=weekday,
                        marks=marks,
                    )
                )

    return rows, warnings


def aggregate_by_weekday(rows: list[CalendarRow]) -> dict[str, dict[str, int]]:
    """学年ごと・曜日ごとの授業日数（〇・○・◯）。"""
    out: dict[str, dict[str, int]] = {g: {wd: 0 for wd in WEEKDAYS_JA} for g in GRADES}
    for row in rows:
        for i, g in enumerate(GRADES):
            if row.marks[i] in CLASS_MARKS:
                out[g][row.weekday] += 1
    return out


def totals(agg: dict[str, dict[str, int]]) -> dict[str, int]:
    return {g: sum(agg[g].values()) for g in GRADES}


def row_date(row: CalendarRow) -> date:
    return date(row.year, row.month, row.day)


def filter_rows_by_date_range(
    rows: list[CalendarRow],
    start: date | None,
    end: date | None,
) -> list[CalendarRow]:
    """開始日・終了日（いずれも含む）で絞り込み。None はその側を制限なし。"""
    out: list[CalendarRow] = []
    for r in rows:
        d = row_date(r)
        if start is not None and d < start:
            continue
        if end is not None and d > end:
            continue
        out.append(r)
    return out


def teaching_days_per_grade_in_range(
    rows: list[CalendarRow],
    start: date | None,
    end: date | None,
) -> dict[str, int]:
    """期間内で各学年列が 〇・○・◯ の日数。"""
    filtered = filter_rows_by_date_range(rows, start, end)
    out = {g: 0 for g in GRADES}
    for r in filtered:
        for i, g in enumerate(GRADES):
            if r.marks[i] in CLASS_MARKS:
                out[g] += 1
    return out


def row_for_date(rows: list[CalendarRow], d: date) -> CalendarRow | None:
    for r in rows:
        if row_date(r) == d:
            return r
    return None


def is_teaching_day_for_grade(rows: list[CalendarRow], d: date, grade: str) -> bool:
    r = row_for_date(rows, d)
    if r is None:
        return False
    i = GRADES.index(grade)
    return r.marks[i] in CLASS_MARKS
