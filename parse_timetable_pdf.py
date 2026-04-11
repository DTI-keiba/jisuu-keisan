"""教員個人の時間割PDFから、担当クラス表記の出現回数（＝コマ数の目安）を数える。"""

from __future__ import annotations

import io
import re
from datetime import date, timedelta
from typing import BinaryIO

import pdfplumber

from parse_calendar_pdf import GRADES, WEEKDAYS_JA, is_teaching_day_for_grade
from parse_calendar_pdf import CalendarRow

# 組番号（1〜40。Excel 出力の １－１１ など全角2桁も含む）
_KUMI_DIGITS = r"[0-9０-９]{1,2}"
# 1年1組 / １年　１１組 など
_RE_NEN_KUMI = re.compile(
    rf"([1-3１-３])\s*年\s*({_KUMI_DIGITS})\s*組",
    re.UNICODE,
)
# 1-1 / １－１ / １－１１（全角ハイフン）／ 1・2 の誤取りを防ぐ
_RE_HYPHEN = re.compile(
    rf"([1-3１-３])\s*[-－・ｰ]\s*({_KUMI_DIGITS})(?![0-9０-９])",
    re.UNICODE,
)

_MAX_KUMI = 40


def _to_ascii_digit(ch: str) -> str:
    m = {
        "０": "0",
        "１": "1",
        "２": "2",
        "３": "3",
        "４": "4",
        "５": "5",
        "６": "6",
        "７": "7",
        "８": "8",
        "９": "9",
    }
    return m.get(ch, ch)


def _norm_num(s: str) -> str:
    return "".join(_to_ascii_digit(c) for c in s)


def _canon_year(y: str) -> str:
    y = _norm_num(y)
    return y if y in {"1", "2", "3"} else y


def _canon_kumi(k: str) -> str:
    k = _norm_num(k)
    return k.lstrip("0") or "0"


def _kumi_in_range(k: str) -> bool:
    k2 = _norm_num(k)
    if not k2.isdigit():
        return False
    n = int(k2)
    return 1 <= n <= _MAX_KUMI


def canonical_label(year: str, kumi: str) -> str:
    return f"{_canon_year(year)}年{_canon_kumi(kumi)}組"


# 表示・集計キー用（クラスと科目を分ける）
SLOT_KEY_SEP = "／"

# 組表記がないセル用（総合・LHR のみなど）
NO_CLASS_PLACEHOLDER = "—"


def _weekday_header_only(seg: str) -> bool:
    """セル全体が曜日見出しだけなら True（集計対象外）。"""
    t = re.sub(r"[ \u3000\n]+", "", str(seg).strip())
    if not t:
        return False
    if len(t) == 1 and t in WEEKDAYS_JA:
        return True
    for wd in WEEKDAYS_JA:
        if t in (f"{wd}曜", f"{wd}曜日"):
            return True
    return False


def _period_header_only(seg: str) -> bool:
    """セル全体が時限・コマ番号の見出しだけなら True（集計対象外）。"""
    if _RE_NEN_KUMI.search(seg) or _RE_HYPHEN.search(seg):
        return False
    raw = str(seg).strip()
    if "年" in raw or "組" in raw:
        return False
    comp = re.sub(r"[ \u3000\n]+", "", raw)
    comp = "".join(_to_ascii_digit(c) if c in "０１２３４５６７８９" else c for c in comp)
    if re.fullmatch(r"[0-9]{1,2}", comp):
        n = int(comp)
        return 1 <= n <= 15
    m = re.fullmatch(r"([0-9]{1,2})(限|コマ|時限)", comp)
    if m:
        n = int(m.group(1))
        return 1 <= n <= 15
    m = re.fullmatch(r"第([0-9]{1,2})(限|コマ|時限)", comp)
    if m:
        n = int(m.group(1))
        return 1 <= n <= 15
    return False


def is_timetable_axis_label_only(seg: str) -> bool:
    """曜日見出し・何コマ目の数字だけのマスは集計に含めない。"""
    return _weekday_header_only(seg) or _period_header_only(seg)


def slot_keys_from_cell(seg: str, extra_labels: list[str] | None = None) -> list[str]:
    """
    マス1つ分の文字列から、区別して数えるスロットキーを返す。
    - クラス＋科目: 「1年1組／英コミュ」「1年1組／総合」「1年1組／LHR」
    - ※のみ: 休みのマスなので **カウントしない**（空リスト）
    - 組なしで「総合」「LHR」のみ: 「—／総合」「—／LHR」
    """
    extra_labels = extra_labels or []
    seg = str(seg).replace("\n", " ").strip()
    seg = re.sub(r"[ \u3000]+", " ", seg)
    if not seg:
        return []

    if is_timetable_axis_label_only(seg):
        return []

    compact = re.sub(r"\s+", "", seg)
    if compact and all(c in "※＊*" for c in compact):
        return []

    spans = _iter_class_spans(seg)
    if not spans:
        if compact == "総合" or seg.strip() == "総合":
            return [f"{NO_CLASS_PLACEHOLDER}{SLOT_KEY_SEP}総合"]
        if compact.upper() == "LHR" or seg.strip().upper() == "LHR":
            return [f"{NO_CLASS_PLACEHOLDER}{SLOT_KEY_SEP}LHR"]
        for x in sorted(extra_labels, key=len, reverse=True):
            if x and x in seg:
                rest = seg.replace(x, "", 1).strip()
                if rest:
                    return [f"{x}{SLOT_KEY_SEP}{rest}"]
                return [x]
        return [seg]

    if len(spans) == 1:
        s0, e0, lab = spans[0]
        before, after = seg[:s0].strip(), seg[e0:].strip()
        subject = re.sub(r"\s+", " ", f"{before} {after}".strip())
        if subject:
            if subject.upper() == "LHR":
                subject = "LHR"
            return [f"{lab}{SLOT_KEY_SEP}{subject}"]
        return [lab]

    return [lab for _, _, lab in spans]


def split_slot_key_for_display(key: str) -> tuple[str, str]:
    """(クラス側, 科目側)。"""
    if SLOT_KEY_SEP in key:
        a, b = key.split(SLOT_KEY_SEP, 1)
        return a.strip(), b.strip()
    return key, ""


def slot_key_sort_key(key: str) -> tuple:
    """
    集計表の並び用キー。
    1年1組→1年2組→…→2年…→3年…、続いてその他のクラス表記（文字列順）、
    最後に組なし（—）。同一クラス内は科目・内容の文字列順。
    """
    c_part, s_part = split_slot_key_for_display(key)
    c = c_part.strip()
    subj = s_part.strip()
    compact = re.sub(r"\s+", "", c)
    compact = "".join(_to_ascii_digit(ch) if ch in "０１２３４５６７８９" else ch for ch in compact)
    m = re.fullmatch(r"([123])年(\d+)組", compact)
    if m:
        y, ku = int(m.group(1)), int(m.group(2))
        return (0, y, ku, subj, key)
    if c == NO_CLASS_PLACEHOLDER:
        return (2, 0, 0, subj, key)
    return (1, c, subj, key)


def detect_slot_labels_from_segments(
    segments: list[str],
    extra_labels: list[str] | None = None,
) -> list[str]:
    seen: dict[str, None] = {}
    for seg in segments:
        for k in slot_keys_from_cell(seg, extra_labels):
            seen.setdefault(k, None)
    return sorted(seen.keys())


def count_slots_in_segments(
    segments: list[str],
    extra_labels: list[str] | None = None,
) -> dict[str, int]:
    counts: dict[str, int] = {}
    for seg in segments:
        for k in slot_keys_from_cell(seg, extra_labels):
            counts[k] = counts.get(k, 0) + 1
    return counts


def _iter_class_spans(text: str) -> list[tuple[int, int, str]]:
    """文字列内の (start, end, canonical_label)。"""
    spans: list[tuple[int, int, str]] = []
    for rx in (_RE_NEN_KUMI, _RE_HYPHEN):
        for m in rx.finditer(text):
            y, k = m.group(1), m.group(2)
            if not _kumi_in_range(k):
                continue
            if _canon_year(y) not in {"1", "2", "3"}:
                continue
            lab = canonical_label(y, k)
            spans.append((m.start(), m.end(), lab))
    spans.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    # 重なり除去（長いほう優先）
    kept: list[tuple[int, int, str]] = []
    for s, e, lab in spans:
        if any(not (e <= ks or s >= ke) for ks, ke, _ in kept):
            continue
        kept.append((s, e, lab))
    return kept


def _cells_from_tables(page: pdfplumber.page.Page) -> list[str]:
    out: list[str] = []
    for table in page.extract_tables() or []:
        for row in table:
            if not row:
                continue
            for cell in row:
                if cell is None:
                    continue
                t = str(cell).replace("\n", " ").strip()
                if t:
                    out.append(t)
    return out


def _lines_from_words(page: pdfplumber.page.Page) -> list[str]:
    words = page.extract_words(use_text_flow=True) or []
    if not words:
        return []
    # おおまかに y で行に分ける
    tol = 3.0
    lines: list[list[tuple[float, str]]] = []
    for w in sorted(words, key=lambda x: (round(x["top"] / tol) * tol, x["x0"])):
        y = round(w["top"] / tol) * tol
        text = (w.get("text") or "").strip()
        if not text:
            continue
        placed = False
        for row in lines:
            if abs(row[0][0] - y) <= tol * 1.5:
                row.append((w["x0"], text))
                placed = True
                break
        if not placed:
            lines.append([(y, text)])
    out: list[str] = []
    for row in lines:
        row.sort(key=lambda x: x[0])
        out.append(" ".join(t for _, t in row))
    return out


def _ocr_with_tesseract(images: list[bytes]) -> list[str]:
    import pytesseract
    from PIL import Image

    lines: list[str] = []
    for png in images:
        im = Image.open(io.BytesIO(png))
        t = pytesseract.image_to_string(im, lang="jpn+jpn_vert") or ""
        for ln in t.splitlines():
            ln = ln.strip()
            if ln:
                lines.append(ln)
    return lines


def _render_pages_png(pdf_bytes: bytes, dpi: int = 200) -> list[bytes]:
    import fitz

    out: list[bytes] = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        for i in range(len(doc)):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out.append(pix.tobytes("png"))
    finally:
        doc.close()
    return out


def extract_timetable_segments(
    file: BinaryIO | bytes,
    *,
    ocr: bool = False,
) -> tuple[list[str], str]:
    """
    PDF から「コマ1つ分の文字列」の候補リストを返す。
    戻り値: (segments, source_description)
    """
    raw = file if isinstance(file, bytes) else file.read()

    if ocr:
        pngs = _render_pages_png(raw)
        segments = _ocr_with_tesseract(pngs)
        return segments, "ocr(tesseract)"

    segments: list[str] = []
    source = "pdfplumber"

    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        total_chars = 0
        for page in pdf.pages:
            cells = _cells_from_tables(page)
            if cells:
                segments.extend(cells)
                total_chars += sum(len(s) for s in cells)
            else:
                lines = _lines_from_words(page)
                segments.extend(lines)
                total_chars += sum(len(s) for s in lines)
            txt = page.extract_text() or ""
            total_chars += len(txt.strip())

    # 表も単語も薄い → OCR を自動で試す（失敗時は pdfplumber のまま）
    if len(segments) < 8 and total_chars < 80:
        try:
            pngs = _render_pages_png(raw)
            segments = _ocr_with_tesseract(pngs)
            source = "ocr(tesseract・自動)"
        except Exception:
            source = "pdfplumber（文字ほぼなし。OCRオンか Tesseract 導入を試してください）"

    return segments, source


def detect_auto_class_labels(segments: list[str]) -> list[str]:
    seen: dict[str, None] = {}
    for seg in segments:
        for _, _, lab in _iter_class_spans(seg):
            seen.setdefault(lab, None)
    return sorted(seen.keys())


def count_classes(
    segments: list[str],
    class_labels: list[str],
    *,
    match_substring: bool = False,
) -> dict[str, int]:
    """
    class_labels ごとにカウント。
    match_substring=False: N年M組パターンのみ（自動検出ラベル向け）。
    match_substring=True: ユーザー指定ラベルを部分一致（1セグメントにつき最大1回）。
    """
    counts = {lab: 0 for lab in class_labels}
    if not class_labels:
        return counts

    if match_substring:
        for seg in segments:
            for lab in class_labels:
                if lab and lab in seg:
                    counts[lab] += 1
        return counts

    for seg in segments:
        found = {lab for _, _, lab in _iter_class_spans(seg)}
        for lab in found:
            if lab in counts:
                counts[lab] += 1
    return counts


def count_pattern_based(segments: list[str]) -> dict[str, int]:
    """パターン（N年M組・N-M）に一致するものをすべて集計。"""
    counts: dict[str, int] = {}
    for seg in segments:
        for _, _, lab in _iter_class_spans(seg):
            counts[lab] = counts.get(lab, 0) + 1
    return dict(sorted(counts.items(), key=lambda x: (-x[1], x[0])))


def _cell_to_weekday(cell: object) -> str | None:
    if cell is None:
        return None
    t = str(cell).replace("\n", "").replace(" ", "").strip()
    if not t:
        return None
    if len(t) == 1 and t in WEEKDAYS_JA:
        return t
    for wd in WEEKDAYS_JA:
        if t == wd or t.startswith(wd + "曜"):
            return wd
    return None


def _accumulate_grid_from_table(
    table: list[list],
    extra_labels: list[str],
) -> dict[tuple[str, str], int]:
    """(スロットキー（クラス／科目 等）, 曜) -> その曜のコマ数。"""
    counts: dict[tuple[str, str], int] = {}
    if not table or len(table) < 2:
        return counts

    header_row_idx: int | None = None
    wd_cols: dict[int, str] = {}
    for ri, row in enumerate(table[:22]):
        if not row:
            continue
        cols_hit: dict[int, str] = {}
        for ci, cell in enumerate(row):
            wd = _cell_to_weekday(cell)
            if wd:
                cols_hit[ci] = wd
        if len(cols_hit) >= 3:
            header_row_idx = ri
            wd_cols = cols_hit
            break

    if header_row_idx is not None:
        for ri in range(header_row_idx + 1, len(table)):
            row = table[ri]
            if not row:
                continue
            for ci, wd in wd_cols.items():
                if ci >= len(row):
                    continue
                cell = row[ci]
                if cell is None:
                    continue
                seg = str(cell).replace("\n", " ").strip()
                if not seg:
                    continue
                if is_timetable_axis_label_only(seg):
                    continue
                for key in slot_keys_from_cell(seg, extra_labels):
                    counts[(key, wd)] = counts.get((key, wd), 0) + 1
        return counts

    wd_rows: dict[int, str] = {}
    for ri, row in enumerate(table[:32]):
        if not row:
            continue
        wd = _cell_to_weekday(row[0])
        if wd:
            wd_rows[ri] = wd
    if len(wd_rows) >= 3:
        for ri, wd in wd_rows.items():
            row = table[ri]
            for ci in range(1, len(row)):
                cell = row[ci]
                if cell is None:
                    continue
                seg = str(cell).replace("\n", " ").strip()
                if not seg:
                    continue
                if is_timetable_axis_label_only(seg):
                    continue
                for key in slot_keys_from_cell(seg, extra_labels):
                    counts[(key, wd)] = counts.get((key, wd), 0) + 1
    return counts


def aggregate_weekly_grid_to_class_totals(
    weekly_grid: dict[tuple[str, str], int] | None,
) -> dict[str, int]:
    """(スロットキー, 曜) の内訳をスロット別コマ数に合算。"""
    if not weekly_grid:
        return {}
    out: dict[str, int] = {}
    for (cls, _), n in weekly_grid.items():
        out[cls] = out.get(cls, 0) + n
    return out


def extract_weekly_class_weekday_counts(
    pdf_bytes: bytes,
    *,
    ocr: bool = False,
    extra_labels: list[str] | None = None,
) -> dict[tuple[str, str], int] | None:
    """
    表の見出しに 月火水木金土日 が並ぶ時間割から、(クラス, 曜) ごとの週あたりコマ数。
    OCR のみのときは None。
    """
    if ocr:
        return None
    extra_labels = extra_labels or []
    merged: dict[tuple[str, str], int] = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                part = _accumulate_grid_from_table(table, extra_labels)
                for k, v in part.items():
                    merged[k] = merged.get(k, 0) + v
    return merged if merged else None


def parse_label_grade_overrides(text: str) -> dict[str, str]:
    """「特進=2年」のような行をパース。値は 1年/2年/2国/3年 のいずれか。"""
    out: dict[str, str] = {}
    for ln in text.splitlines():
        ln = ln.strip()
        if not ln or "=" not in ln:
            continue
        a, b = ln.split("=", 1)
        k, v = a.strip(), b.strip()
        if k and v in GRADES:
            out[k] = v
    return out


def infer_track_for_class_label(
    label: str,
    overrides: dict[str, str] | None = None,
) -> str | None:
    """クラス表記から行事予定の学年列（1年・2年・2国・3年）へ。"""
    overrides = overrides or {}
    if label in overrides:
        v = overrides[label]
        return v if v in GRADES else None
    m = re.match(r"^([123])年", label)
    if m:
        return {"1": "1年", "2": "2年", "3": "3年"}[m.group(1)]
    if "2国" in label:
        return "2国"
    return None


def infer_track_for_slot_key(
    slot_key: str,
    overrides: dict[str, str] | None = None,
) -> str | None:
    """スロットキーから行事予定の学年列へ。組なし（—／…）は None（いずれか学年が〇の日で按分）。"""
    overrides = overrides or {}
    if slot_key in overrides:
        v = overrides[slot_key]
        return v if v in GRADES else None
    head = slot_key.split(SLOT_KEY_SEP, 1)[0].strip()
    if head in (NO_CLASS_PLACEHOLDER, "―", "－", "-"):
        return None
    return infer_track_for_class_label(head, overrides)


def _any_grade_teaching_day(rows: list[CalendarRow], d: date) -> bool:
    return any(is_teaching_day_for_grade(rows, d, g) for g in GRADES)


def _daterange(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def _ja_weekday(d: date) -> str:
    return WEEKDAYS_JA[d.weekday()]


def project_lessons_in_period(
    *,
    weekly_class_weekday: dict[tuple[str, str], int] | None,
    weekly_flat_counts: dict[str, int],
    calendar_rows: list[CalendarRow],
    start: date,
    end: date,
    class_labels: list[str],
    label_grade_overrides: dict[str, str] | None = None,
) -> tuple[dict[str, int], str]:
    """
    行事予定の〇日と突合し、期間内の担当コマ数を推定。
    戻り値: (クラス別コマ数, 算出方法の説明)
    """
    label_grade_overrides = label_grade_overrides or {}
    if start > end:
        return {}, "開始日が終了日より後です"

    out: dict[str, int] = {c: 0 for c in class_labels}

    if weekly_class_weekday:
        for d in _daterange(start, end):
            wd = _ja_weekday(d)
            for cls in class_labels:
                n = weekly_class_weekday.get((cls, wd), 0)
                if n == 0:
                    continue
                track = infer_track_for_slot_key(cls, label_grade_overrides)
                if track is None:
                    if _any_grade_teaching_day(calendar_rows, d):
                        out[cls] += n
                elif is_teaching_day_for_grade(calendar_rows, d, track):
                    out[cls] += n
        return (
            out,
            "表から読み取った「曜日×（クラス／科目）」のコマを、行事予定の〇日と日付ごとに突合"
            "（組が「—」の行は、いずれかの学年が〇の日にカウント）。",
        )

    teach = {g: 0 for g in GRADES}
    teach_any = 0
    for d in _daterange(start, end):
        if _any_grade_teaching_day(calendar_rows, d):
            teach_any += 1
        for g in GRADES:
            if is_teaching_day_for_grade(calendar_rows, d, g):
                teach[g] += 1

    note = (
        "概算: 曜日列を表から判別できなかったため、"
        "「週あたりコマ数×（期間内の該当学年の授業日数）÷5」で按分しています（平日に均等と仮定）。"
        " 組が「—」の行は「いずれかの学年が〇の日数」で按分。"
    )
    for cls in class_labels:
        w = weekly_flat_counts.get(cls, 0)
        if w == 0:
            continue
        track = infer_track_for_slot_key(cls, label_grade_overrides)
        if track is None:
            out[cls] = int(round(teach_any * w / 5.0))
        else:
            ng = teach[track]
            out[cls] = int(round(ng * w / 5.0))
    return out, note
