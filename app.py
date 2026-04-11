"""行事予定PDF・教員時間割PDFの集計（Streamlit）。"""

from __future__ import annotations

import io

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
)
from parse_timetable_pdf import (
    aggregate_weekly_grid_to_class_totals,
    count_slots_in_segments,
    extract_timetable_segments,
    extract_weekly_class_weekday_counts,
    parse_label_grade_overrides,
    project_lessons_in_period,
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
        st.rerun()


def _period_bounds() -> tuple:
    """サイドバーの period_range（開始・終了の2日）からタプルを返す。"""
    pr = st.session_state.get("period_range")
    if isinstance(pr, (tuple, list)) and len(pr) == 2:
        return pr[0], pr[1]
    if isinstance(pr, (tuple, list)) and len(pr) == 1:
        return pr[0], pr[0]
    return None, None


st.set_page_config(page_title="行事予定・時間割 集計アプリ", layout="wide")

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

        st.date_input(
            "集計期間（カレンダーで選択）",
            min_value=dmin,
            max_value=dmax,
            key="period_range",
            help="クリックするとカレンダーが開きます。開始日と終了日をタップして範囲を指定してください（キーボード入力は使わずに選べます）。",
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
    up_cal = st.file_uploader("行事予定PDF", type=["pdf"], key="cal_pdf")

    rows = None
    warnings: list[str] = []
    if up_cal is not None:
        rows, warnings = parse_pdf(io.BytesIO(up_cal.read()))
        if rows:
            st.session_state["cal_rows"] = rows
            _apply_new_calendar(rows, up_cal.name)
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

            agg = aggregate_by_weekday(rows_f)
            tot = totals(agg)

            df = pd.DataFrame(
                [
                    {**{"学年": g}, **{wd: agg[g][wd] for wd in WEEKDAYS_JA}, **{"計": tot[g]}}
                    for g in GRADES
                ]
            )
            st.subheader("集計結果（期間内）")
            st.dataframe(df, use_container_width=True, hide_index=True)

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
            st.session_state["cal_rows"] = cr
            _apply_new_calendar(cr, cal_upload_tt.name)
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
        for k in sorted(merged.keys(), key=lambda x: (-merged.get(x, 0), x)):
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
            chart_col = (
                "期間内コマ（推定）"
                if period_projection_ok and "期間内コマ（推定）" in out_df.columns
                else "PDF検出コマ"
            )
            if chart_col in out_df.columns:
                st.bar_chart(out_df.set_index("区分（集計キー）")[[chart_col]])

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
