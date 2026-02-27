import math
from datetime import date, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st

WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

def daterange(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)

def is_sunday(d: date) -> bool:
    # Python: Monday=0 ... Sunday=6
    return d.weekday() == 6

def make_week_blocks(total_blocks: int):
    return list(range(1, total_blocks + 1))

def generate_schedule(
    start_date: date,
    weeks: int,
    ha: float,
    blocks_per_ha: int,
    trees_per_block: int,
    events_per_week: int,
    max_blocks_per_day: int,
    rest_on_sunday: bool = True,
):
    # blocks
    total_blocks = int(round(ha * blocks_per_ha))
    if total_blocks <= 0:
        return pd.DataFrame(), {"total_blocks": 0}

    # horizon date
    end_date = start_date + timedelta(days=weeks * 7 - 1)

    # date list
    all_days = list(daterange(start_date, end_date))
    if rest_on_sunday:
        work_days = [d for d in all_days if not is_sunday(d)]
    else:
        work_days = all_days

    # group by week (Mon-Sun) based on ISO week
    # We'll slice sequentially by 7-day windows starting from start_date for simplicity.
    rows = []
    block_ids = make_week_blocks(total_blocks)

    warnings = []

    for w in range(weeks):
        wk_start = start_date + timedelta(days=w * 7)
        wk_end = wk_start + timedelta(days=6)

        # candidate days this week (in range + not Sunday if enabled)
        week_days = [d for d in daterange(wk_start, wk_end) if (d >= start_date and d <= end_date)]
        if rest_on_sunday:
            week_days = [d for d in week_days if not is_sunday(d)]

        if len(week_days) == 0:
            continue

        # total "slots" in the week
        total_slots = len(week_days) * max_blocks_per_day
        required = total_blocks * events_per_week

        if required > total_slots:
            warnings.append(
                f"Week {w+1}: 必要割当({required}件=ブロック{total_blocks}×{events_per_week}回/週)が、"
                f"週の上限枠({total_slots}件=稼働日{len(week_days)}×{max_blocks_per_day}ブロック/日)を超えています。"
                f" → max_blocks_per_day を増やすか、events_per_week を下げてください。"
            )

        # Build slots list (day repeated by max_blocks_per_day)
        slots = []
        for d in week_days:
            slots.extend([d] * max_blocks_per_day)

        # Assign blocks to slots
        # To spread repeats when events_per_week >= 2, rotate each pass.
        assignments = {d: [] for d in week_days}
        slot_idx = 0

        for k in range(events_per_week):
            rotated = block_ids[k:] + block_ids[:k]  # rotation helps separation
            for b in rotated:
                if slot_idx >= len(slots):
                    break
                day = slots[slot_idx]
                assignments[day].append(b)
                slot_idx += 1

        # Create rows per day
        for d in week_days:
            b_list = assignments.get(d, [])
            rows.append({
                "日付": d.isoformat(),
                "曜日": WEEKDAYS_JA[d.weekday()],
                "週": w + 1,
                "ブロック数": len(b_list),
                "ブロック一覧": ", ".join([str(x) for x in b_list]),
                "本数(推定)": len(b_list) * trees_per_block,
            })

    df = pd.DataFrame(rows)

    meta = {
        "ha": ha,
        "blocks_per_ha": blocks_per_ha,
        "trees_per_block": trees_per_block,
        "total_blocks": total_blocks,
        "events_per_week": events_per_week,
        "max_blocks_per_day": max_blocks_per_day,
        "weeks": weeks,
        "rest_on_sunday": rest_on_sunday,
        "warnings": warnings,
    }
    return df, meta

def to_excel_bytes(df: pd.DataFrame, meta: dict):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        # schedule
        df.to_excel(writer, index=False, sheet_name="Schedule")

        # summary/meta
        summary = pd.DataFrame([
            ["ha", meta["ha"]],
            ["blocks_per_ha", meta["blocks_per_ha"]],
            ["trees_per_block", meta["trees_per_block"]],
            ["total_blocks", meta["total_blocks"]],
            ["events_per_week", meta["events_per_week"]],
            ["max_blocks_per_day", meta["max_blocks_per_day"]],
            ["weeks", meta["weeks"]],
            ["rest_on_sunday", meta["rest_on_sunday"]],
        ], columns=["項目", "値"])
        summary.to_excel(writer, index=False, sheet_name="Meta")

        if meta.get("warnings"):
            warn = pd.DataFrame({"警告": meta["warnings"]})
            warn.to_excel(writer, index=False, sheet_name="Warnings")

    bio.seek(0)
    return bio.getvalue()


# -------------------------
# UI
# -------------------------
st.set_page_config(page_title="水やりスケジュール自動生成", layout="wide")
st.title("水やりスケジュール自動生成（ブロック単位 / 日曜休み）")

with st.sidebar:
    st.header("入力")
    ha = st.number_input("対象面積（ha）", min_value=0.0, value=1.0, step=0.5)
    blocks_per_ha = st.number_input("1haあたりブロック数", min_value=1, value=17, step=1)
    trees_per_block = st.number_input("1ブロックあたり本数", min_value=1, value=117, step=1)

    start_date = st.date_input("開始日", value=date.today())
    weeks = st.number_input("何週間先まで作る？", min_value=1, value=8, step=1)

    events_per_week = st.selectbox("各ブロックを週に何回水やりする？", options=[1, 2, 3], index=0)
    max_blocks_per_day = st.number_input("1日に水やりする最大ブロック数（分割したい上限）", min_value=1, value=10, step=1)

    rest_on_sunday = st.checkbox("日曜は休み（固定）", value=True)

df, meta = generate_schedule(
    start_date=start_date,
    weeks=int(weeks),
    ha=float(ha),
    blocks_per_ha=int(blocks_per_ha),
    trees_per_block=int(trees_per_block),
    events_per_week=int(events_per_week),
    max_blocks_per_day=int(max_blocks_per_day),
    rest_on_sunday=bool(rest_on_sunday),
)

col1, col2 = st.columns([2, 1])

with col2:
    st.subheader("サマリー")
    st.write(f"- 対象面積: **{meta.get('ha', 0)} ha**")
    st.write(f"- 総ブロック数: **{meta.get('total_blocks', 0)} blocks**")
    st.write(f"- 1ブロック本数: **{meta.get('trees_per_block', 0)} 本**")
    st.write(f"- 週あたり回数: **{meta.get('events_per_week', 0)} 回/週**")
    st.write(f"- 1日上限: **{meta.get('max_blocks_per_day', 0)} blocks/日**")
    if meta.get("warnings"):
        st.error("⚠ 容量不足の可能性があります")
        for w in meta["warnings"]:
            st.write(f"- {w}")
    else:
        st.success("OK：週内に割り振り可能です")

with col1:
    st.subheader("生成されたスケジュール")
    if df.empty:
        st.info("haが0、または期間が短すぎる等でスケジュールが生成されていません。")
    else:
        st.dataframe(df, use_container_width=True, height=520)

        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="CSVをダウンロード",
            data=csv_bytes,
            file_name="watering_schedule.csv",
            mime="text/csv",
        )

        xlsx_bytes = to_excel_bytes(df, meta)
        st.download_button(
            label="Excelをダウンロード",
            data=xlsx_bytes,
            file_name="watering_schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
