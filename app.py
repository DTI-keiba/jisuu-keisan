"""行事予定PDF・教員時間割PDFの集計（Streamlit）。"""

from __future__ import annotations

import io
import uuid

import pandas as pd
import streamlit as st

from parse_calendar_pdf import (
    GRADES,
    WEEKDAYS_JA,
    aggregate_by_weekday,
    filter_rows_by_date_range,
    parse_pdf,
    row_date,
    teaching_days_per_grade_in_range,
    totals,
    weekday_teaching_totals,
)
from parse_timetable_pdf import (
    aggregate_weekly_grid_to_class_totals,
    count_slots_in_segments,
    extract_timetable_segments,
    extract_weekly_class_weekday_counts,
    parse_label_grade_overrides,
    project_lessons_in_period,
    slot_key_sort_key,
    split_slot_key_for_display,
)

_MENU_CAL = "行事予定（学年・曜日）"
_MENU_TT = "時間割（担当コマ）"


def _cal_fingerprint(name: str, n_rows: int, d_min, d_max) -> str:
    return f"{name}|{n_rows}|{d_min}|{d_max}"


def _apply_new_calendar(rows: list, source_name: str) -> None:
    if not rows:
        return
    ds = [row_date(r) for r in rows]
    d_min, d_max = min(ds), max(ds)
    fp = _cal_fingerprint(source_name, len(rows), d_min, d_max)
    if st.session_state.get("_cal_fp") != fp:
        st.session_state["_cal_fp"] = fp
        st.session_state["period_range"] = (d_min, d_max)
        st.session_state["period_start"] = d_min
        st.session_state["period_end"] = d_max
        st.rerun()


def _register_calendar_in_library(rows: list, source_name: str) -> None:
    """解析済み行事予定をライブラリに登録し、選択中にする（同一内容は上書き）。"""
    if not rows:
        return
    ds = [row_date(r) for r in rows]
    d_min, d_max = min(ds), max(ds)
    fp = _cal_fingerprint(source_name, len(rows), d_min, d_max)
    lib: list[dict] = st.session_state.setdefault("cal_library", [])
    for e in lib:
        if e.get("fingerprint") == fp:
            e["rows"] = list(rows)
            e["name"] = source_name
            e["d_min"] = d_min
            e["d_max"] = d_max
            e["n"] = len(rows)
            st.session_state["cal_active_id"] = e["id"]
            st.session_state["cal_rows"] = list(rows)
            st.session_state["cal_library_select"] = e["id"]
            _apply_new_calendar(rows, source_name)
            return
    eid = uuid.uuid4().hex[:10]
    lib.append(
        {
            "id": eid,
            "name": source_name,
            "rows": list(rows),
            "fingerprint": fp,
            "d_min": d_min,
            "d_max": d_max,
            "n": len(rows),
        }
    )
    st.session_state["cal_active_id"] = eid
    st.session_state["cal_rows"] = list(rows)
    st.session_state["cal_library_select"] = eid
    _apply_new_calendar(rows, source_name)


def _migrate_cal_library_from_session_rows() -> None:
    """古いセッション（cal_rows のみ）からライブラリを1件だけ作る。"""
    lib: list[dict] = st.session_state.setdefault("cal_library", [])
    if lib:
        return
    rows = st.session_state.get("cal_rows")
    if not rows:
        return
    ds = [row_date(r) for r in rows]
    d_min, d_max = min(ds), max(ds)
    eid = uuid.uuid4().hex[:10]
    fp = _cal_fingerprint("読み込み済み", len(rows), d_min, d_max)
    lib.append(
        {
            "id": eid,
            "name": "読み込み済み",
            "rows": list(rows),
            "fingerprint": fp,
            "d_min": d_min,
            "d_max": d_max,
            "n": len(rows),
        }
    )
    st.session_state["cal_active_id"] = eid
    st.session_state["cal_library_select"] = eid


def _on_cal_library_select_change() -> None:
    eid = st.session_state.get("cal_library_select")
    if not eid:
        return
    for e in st.session_state.get("cal_library") or []:
        if e["id"] == eid:
            st.session_state["cal_rows"] = list(e["rows"])
            _apply_new_calendar(e["rows"], e["name"])
            return


def _period_bounds() -> tuple:
    """サイドバーの開始日・終了日（または period_range）からタプルを返す。"""
    ps = st.session_state.get("period_start")
    pe = st.session_state.get("period_end")
    if ps is not None and pe is not None:
        return ps, pe
    pr = st.session_state.get("period_range")
    if isinstance(pr, (tuple, list)) and len(pr) == 2:
        return pr[0], pr[1]
    if isinstance(pr, (tuple, list)) and len(pr) == 1:
        return pr[0], pr[0]
    return None, None


def _calendar_agg_tables(rows_f: list) -> tuple[pd.DataFrame, pd.DataFrame, dict, dict]:
    """
    期間内の行事予定行から、学年×曜日の表・曜日×学年の表と agg / tot を返す。
    rows_f: filter_rows_by_date_range 済みの CalendarRow のリスト
    """
    agg = aggregate_by_weekday(rows_f)
    tot = totals(agg)
    df_grade_rows = pd.DataFrame(
        [
            {**{"学年": g}, **{wd: agg[g][wd] for wd in WEEKDAYS_JA}, **{"計": tot[g]}}
            for g in GRADES
        ]
    )
    wd_tot = weekday_teaching_totals(agg)
    df_weekday_rows = pd.DataFrame(
        [
            {
                "曜日": wd,
                **{g: agg[g][wd] for g in GRADES},
                "合計（全学年）": wd_tot[wd],
            }
            for wd in WEEKDAYS_JA
        ]
    )
    return df_grade_rows, df_weekday_rows, agg, tot


st.set_page_config(page_title="行事予定・時間割 集計アプリ", layout="wide")

st.session_state.setdefault("cal_library", [])
st.session_state.setdefault("cal_active_id", None)
_migrate_cal_library_from_session_rows()

cal_rows = list(st.session_state.get("cal_rows") or [])

with st.sidebar:
    st.header("メニュー")
    page = st.radio("画面を選ぶ", [_MENU_CAL, _MENU_TT], key="app_nav")
    st.caption("PDFを読み取り、ここで集計します。")
    st.divider()

    st.subheader("集計期間")
    if cal_rows:
        cds = [row_date(r) for r in cal_rows]
        dmin, dmax = min(cds), max(cds)
        if "period_range" not in st.session_state:
            st.session_state.period_range = (dmin, dmax)
        pr = st.session_state.period_range
        if (
            not isinstance(pr, (tuple, list))
            or len(pr) != 2
            or pr[0] is None
            or pr[1] is None
        ):
            st.session_state.period_range = (dmin, dmax)
            pr = st.session_state.period_range
        ps, pe = pr[0], pr[1]
        ps = max(dmin, min(ps, dmax))
        pe = max(dmin, min(pe, dmax))
        if ps > pe:
            ps, pe = dmin, dmax
        if (ps, pe) != tuple(pr):
            st.session_state.period_range = (ps, pe)
            pr = st.session_state.period_range

        if "period_start" not in st.session_state:
            st.session_state.period_start = pr[0]
        if "period_end" not in st.session_state:
            st.session_state.period_end = pr[1]

        ps = max(dmin, min(st.session_state.period_start, dmax))
        pe = max(dmin, min(st.session_state.period_end, dmax))
        if ps > pe:
            ps, pe = dmin, dmax
        if ps != st.session_state.period_start:
            st.session_state.period_start = ps
        if pe != st.session_state.period_end:
            st.session_state.period_end = pe
        st.session_state.period_range = (st.session_state.period_start, st.session_state.period_end)

        tab_s, tab_e = st.tabs(["開始日", "終了日"])
        with tab_s:
            st.date_input(
                "開始日",
                min_value=dmin,
                max_value=dmax,
                key="period_start",
                help="行事予定PDFの範囲内で、集計のはじめの日を選びます。",
            )
        with tab_e:
            st.date_input(
                "終了日",
                min_value=dmin,
                max_value=dmax,
                key="period_end",
                help="行事予定PDFの範囲内で、集計の終わりの日を選びます。",
            )
        st.session_state.period_range = (
            st.session_state.period_start,
            st.session_state.period_end,
        )
        st.caption(f"行事予定PDFの全日付: **{dmin}** ～ **{dmax}**")
        st.caption("行事予定・時間割の期間集計の両方に使われます。")
    else:
        st.caption("行事予定PDFを読み込むと、ここで期間を設定できます。")

    st.divider()
    st.markdown("**起動**  \n`streamlit run app.py`")

st.title("行事予定・時間割 集計アプリ")
st.caption("ブラウザでPDFをアップロードして利用します（スプレッドシート単体では同じPDFは読み取れません）。")

start_d, end_d = _period_bounds()

# ----- 行事予定 -----
if page == _MENU_CAL:
    st.caption("1年・2年・2国・3年列の 〇・○・◯ を授業日として、暦に沿って学年ごと・曜日ごとに集計します。")
    st.info("**集計期間**は左サイドバーで変更できます。", icon="📅")

    cal_lib: list[dict] = list(st.session_state.get("cal_library") or [])
    if cal_lib:
        ids = [e["id"] for e in cal_lib]
        labels = {e["id"]: f"{e['name']}（{e['d_min']}～{e['d_max']}・{e['n']}日）" for e in cal_lib}
        if st.session_state.get("cal_library_select") not in ids:
            st.session_state.cal_library_select = (
                st.session_state.cal_active_id
                if st.session_state.cal_active_id in ids
                else ids[0]
            )
        default_ix = ids.index(st.session_state.cal_library_select)
        st.selectbox(
            "保存済みの行事予定（このセッション内）",
            options=ids,
            index=default_ix,
            format_func=lambda x: labels.get(x, x),
            key="cal_library_select",
            on_change=_on_cal_library_select_change,
            help="PDFを読み取った一覧から選ぶと、その内容で集計します。ブラウザを閉じると一覧は消えます。",
        )

    up_cal = st.file_uploader(
        "行事予定PDF（新規追加・同じ内容は上書き）",
        type=["pdf"],
        key="cal_pdf",
    )

    rows = None
    warnings: list[str] = []
    if up_cal is not None:
        up_cal.seek(0)
        rows, warnings = parse_pdf(io.BytesIO(up_cal.read()))
        if rows:
            _register_calendar_in_library(rows, up_cal.name)
        rows = list(st.session_state.get("cal_rows") or [])
    elif cal_rows:
        rows = cal_rows

    if up_cal is None and not cal_rows:
        st.info("教員用行事予定などのPDFをアップロードしてください。")
    else:
        if warnings:
            with st.expander("警告・メモ（曜日不一致など）", expanded=False):
                for w in warnings:
                    st.text(w)

        if not rows:
            st.error("集計できる行がありません。PDFの形式が想定と異なる可能性があります。")
        elif start_d is None or end_d is None:
            st.warning("サイドバーで集計期間を確定してください。")
        else:
            if start_d > end_d:
                st.error("開始日は終了日以前にしてください（サイドバーで修正）。")
                st.stop()

            rows_f = filter_rows_by_date_range(rows, start_d, end_d)
            if not rows_f:
                st.warning("この期間に該当する日がありません。")
                st.stop()

            teach_n = teaching_days_per_grade_in_range(rows, start_d, end_d)
            st.caption(
                f"期間: **{start_d}** ～ **{end_d}**（{len(rows_f)} 日分） / "
                + " ".join(f"{g}の授業日 {teach_n[g]}日" for g in GRADES)
            )

            df, df_wd, agg, tot = _calendar_agg_tables(rows_f)

            st.subheader("学年ごとの曜日別授業数（期間内）")
            st.caption(
                "**行が学年**、**列が月〜日**です。セルはその学年がその曜日に **〇・○・◯ だった回数**。"
                " 列「計」は期間内のその学年の合計です。"
            )
            st.dataframe(df, use_container_width=True, hide_index=True)

            st.subheader("曜日別の授業数（学年別内訳・期間内）")
            st.caption(
                "**行が曜日**、**列が学年**です（上の表と同じ数値の見せ方違い）。"
                "「合計（全学年）」は4学年の足し算です。"
            )
            st.dataframe(df_wd, use_container_width=True, hide_index=True)

            chart_df = pd.DataFrame(
                {g: [agg[g][wd] for wd in WEEKDAYS_JA] for g in GRADES},
                index=list(WEEKDAYS_JA),
            )
            st.subheader("曜日別（積み上げ）")
            st.bar_chart(chart_df)

            st.subheader("学年別の合計（期間内の授業日数）")
            st.bar_chart(
                pd.DataFrame({"授業日数": [tot[g] for g in GRADES]}, index=list(GRADES))
            )

            with st.expander(f"解析した行数（全期間）: {len(rows)} 行", expanded=False):
                sample = pd.DataFrame(
                    [
                        {
                            "日付": f"{r.year}-{r.month:02d}-{r.day:02d}",
                            "曜": r.weekday,
                            **dict(zip(GRADES, r.marks)),
                        }
                        for r in rows[:200]
                    ]
                )
                st.dataframe(sample, use_container_width=True, hide_index=True)
                if len(rows) > 200:
                    st.caption("先頭200行のみ表示")

            csv_buf = io.StringIO()
            df.to_csv(csv_buf, index=False, encoding="utf-8-sig")
            st.download_button(
                "CSVをダウンロード（期間内集計）",
                data=csv_buf.getvalue().encode("utf-8-sig"),
                file_name="jukunichi_by_weekday_range.csv",
                mime="text/csv",
                key="dl_cal",
            )
            csv_wd = io.StringIO()
            df_wd.to_csv(csv_wd, index=False, encoding="utf-8-sig")
            st.download_button(
                "CSVをダウンロード（曜日別・学年別）",
                data=csv_wd.getvalue().encode("utf-8-sig"),
                file_name="jukunichi_weekday_by_grade.csv",
                mime="text/csv",
                key="dl_cal_wd",
            )

# ----- 時間割 -----
elif page == _MENU_TT:
    st.caption(
        "個人の時間割PDFからマスを読み取り、**クラス＋科目**が違えば別項目として数えます（例: 1年1組／英コミュ、1年1組／総合、1年1組／LHR）。"
        " **※だけのマスは休みのためカウントしません。** 総合・LHRも **組表記があれば同じ行にクラス** を付けます（組だけないときはクラス欄に「—」）。"
        " **期間集計は左サイドバーの「集計期間」を使います**（行事予定を読み込んだあと）。"
    )
    if cal_rows:
        st.info(
            f"現在の集計期間（サイドバーと共通）: **{start_d}** ～ **{end_d}**",
            icon="📅",
        )
    use_ocr = st.checkbox(
        "OCRを使う（スキャン・画像だけのPDF。Tesseract + 日本語データが必要）",
        value=False,
        key="tt_ocr",
    )
    if use_ocr:
        st.info(
            "Windows の例: [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) をインストールし、"
            "`PATH` に `tesseract.exe` を通すか、環境変数で設定してください。"
        )

    up_tt = st.file_uploader("時間割PDF", type=["pdf"], key="tt_pdf")

    cal_upload_tt = st.file_uploader(
        "行事予定PDF（省略可・行事予定画面と同じデータに追加）",
        type=["pdf"],
        key="tt_cal_pdf",
        help="「行事予定」画面で既に読み込んでいれば不要です。",
    )
    if cal_upload_tt is not None:
        cr, wcal = parse_pdf(io.BytesIO(cal_upload_tt.read()))
        if wcal:
            with st.expander("行事予定PDFの警告", expanded=False):
                for w in wcal:
                    st.text(w)
        if cr:
            _register_calendar_in_library(cr, cal_upload_tt.name)
        elif not st.session_state.get("cal_rows"):
            st.warning("行事予定PDFを解釈できませんでした。")

    extra = st.text_area(
        "追加の担当クラス名（1行に1つ。表に出る表記と同じ文字列。部分一致で1セルにつき1回カウント）",
        placeholder="例:\n2国\n特進\n理数",
        height=100,
        key="tt_extra",
    )
    extra_labels = [ln.strip() for ln in extra.splitlines() if ln.strip()]

    overrides_txt = st.text_area(
        "学年コースの上書き（期間集計用・任意。1行に「ラベル=1年」の形式）",
        placeholder="例:\n特進=2年\n理数=2年",
        height=80,
        key="tt_override",
    )
    label_grade_overrides = parse_label_grade_overrides(overrides_txt)

    pasted = st.text_area(
        "またはテキストを貼り付け（PDFが読めないとき。1行＝1セル相当）",
        height=80,
        key="tt_paste",
    )

    cal_rows = list(st.session_state.get("cal_rows") or [])
    p_start, p_end = start_d, end_d
    if not cal_rows:
        st.info("期間内のコマ数を出すには、行事予定PDFを読み込んでください（左「行事予定」または上のアップロード）。")
    elif p_start is None or p_end is None:
        st.info("サイドバーで集計期間を設定してください。")

    if up_tt is None and not pasted.strip():
        st.info("時間割のPDFをアップロードするか、テキストを貼り付けてください。")
    else:
        segments: list[str] = []
        source = ""
        pdf_bytes: bytes | None = None

        if pasted.strip():
            segments.extend(ln.strip() for ln in pasted.splitlines() if ln.strip())
            source = "貼り付けテキスト"

        if up_tt is not None:
            pdf_bytes = up_tt.read()
            try:
                segs, src = extract_timetable_segments(pdf_bytes, ocr=use_ocr)
                source = src if not segments else f"{source} + {src}"
                segments.extend(segs)
            except Exception as e:
                st.error(f"PDFの読み取りに失敗しました: {e}")
                st.stop()

        if not segments:
            st.warning("本文が取得できませんでした。OCRを有効にするか、テキスト貼り付けを試してください。")
            st.stop()

        weekly_grid: dict[tuple[str, str], int] | None = None
        if pdf_bytes is not None:
            weekly_grid = extract_weekly_class_weekday_counts(
                pdf_bytes, ocr=use_ocr, extra_labels=extra_labels
            )

        grid_totals = aggregate_weekly_grid_to_class_totals(weekly_grid)

        slot_seg_counts = count_slots_in_segments(segments, extra_labels)
        known_keys = set(grid_totals.keys()) | set(slot_seg_counts.keys())
        counts_user_extra: dict[str, int] = {}
        for lab in extra_labels:
            if not lab or any(lab in k for k in known_keys):
                continue
            n = sum(1 for seg in segments if lab in seg)
            if n:
                counts_user_extra[lab] = n

        all_for_table = sorted(
            set(grid_totals.keys()) | set(slot_seg_counts.keys()) | set(counts_user_extra.keys())
        )

        merged: dict[str, int] = {}
        for lab in all_for_table:
            if grid_totals and lab in grid_totals:
                base = grid_totals[lab]
            else:
                base = slot_seg_counts.get(lab, 0)
            merged[lab] = base + counts_user_extra.get(lab, 0)

        slot_ref_counts = slot_seg_counts

        period_note = ""
        period_counts: dict[str, int] = {}
        period_projection_ok = (
            bool(cal_rows)
            and p_start is not None
            and p_end is not None
            and p_start <= p_end
        )
        if cal_rows and p_start is not None and p_end is not None:
            if p_start > p_end:
                st.error("期間の開始日は終了日以前にしてください（サイドバーで修正）。")
            else:
                period_counts, period_note = project_lessons_in_period(
                    weekly_class_weekday=weekly_grid,
                    weekly_flat_counts=merged,
                    calendar_rows=cal_rows,
                    start=p_start,
                    end=p_end,
                    class_labels=all_for_table,
                    label_grade_overrides=label_grade_overrides,
                )

        st.subheader("担当コマ数（クラス・科目など別）")
        if source:
            st.caption(f"取得方法: {source} / セグメント数: {len(segments)}")
        if weekly_grid and pdf_bytes is not None and not use_ocr:
            st.caption(
                "**テキスト／表データのPDF**として読み取っています。"
                "スキャン画像ではないため、通常は **OCR をオンにする必要はありません**。"
            )
        if weekly_grid:
            st.caption(
                "時間割表から **曜日の行または列** を検出しました。"
                "コマ数は表のマスを優先し、**科目名が異なれば別行**にします。期間集計は行事予定の〇日と突合します。"
            )
        elif pdf_bytes is not None and not use_ocr:
            st.caption("時間割から曜日見出しを検出できなかった場合、期間集計は **概算（÷5按分）** になります。")

        rows_out: list[dict] = []
        for k in sorted(merged.keys(), key=slot_key_sort_key):
            c_part, s_part = split_slot_key_for_display(k)
            rowd: dict = {
                "区分（集計キー）": k,
                "クラス": c_part,
                "科目・内容": s_part,
                "PDF検出コマ": merged[k],
            }
            if period_projection_ok:
                rowd["期間内コマ（推定）"] = period_counts.get(k, 0)
            rows_out.append(rowd)

        out_df = pd.DataFrame(rows_out)
        if out_df.empty:
            st.warning(
                "時間割からマスを検出できませんでした。"
                "「追加の担当クラス名」に表記を足すか、OCR・テキスト貼り付けを試してください。"
            )
            if slot_ref_counts:
                st.write("セグメントから検出した区分（参考）:")
                st.dataframe(
                    pd.DataFrame(
                        [{"区分": k, "回数": v} for k, v in slot_ref_counts.items()]
                    ),
                    hide_index=True,
                    use_container_width=True,
                )
        else:
            st.dataframe(out_df, use_container_width=True, hide_index=True)

            if (
                cal_rows
                and p_start is not None
                and p_end is not None
                and p_start <= p_end
            ):
                rows_cal_f = filter_rows_by_date_range(cal_rows, p_start, p_end)
                if rows_cal_f:
                    st.subheader("学年ごとの曜日別授業数（行事予定・期間内）")
                    st.caption(
                        "左の行事予定PDFの〇・○・◯を、**学年×曜日**で数えたものです（サイドバーの期間）。"
                    )
                    df_g, df_wd, _, _ = _calendar_agg_tables(rows_cal_f)
                    st.markdown("**行＝学年、列＝曜日**")
                    st.dataframe(df_g, use_container_width=True, hide_index=True)
                    st.markdown("**行＝曜日、列＝学年**")
                    st.dataframe(df_wd, use_container_width=True, hide_index=True)

        if period_note:
            st.info(period_note)

        csv_tt = io.StringIO()
        if not out_df.empty:
            out_df.to_csv(csv_tt, index=False, encoding="utf-8-sig")
            st.download_button(
                "CSVをダウンロード",
                data=csv_tt.getvalue().encode("utf-8-sig"),
                file_name="timetable_class_counts.csv",
                mime="text/csv",
                key="dl_tt",
            )

        with st.expander("抽出したセル・行のプレビュー（先頭80件）", expanded=False):
            prev = pd.DataFrame({"テキスト": segments[:80]})
            st.dataframe(prev, use_container_width=True, hide_index=True)
