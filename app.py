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


def make_block_ids(total_blocks: int):
    return list(range(1, total_blocks + 1))


def _even_targets(days, total_assignable, max_per_day):
    """
    days: list[date]
    total_assignable: 週の中で実際に割り当て可能な総件数（cap超えならcapまで）
    max_per_day: 1日あたり上限
    戻り値: dict(date -> target)
    """
    D = len(days)
    if D <= 0:
        return {}
    base = total_assignable // D
    rem = total_assignable % D

    targets = {}
    for i, d in enumerate(days):
        t = base + (1 if i < rem else 0)
        targets[d] = min(t, max_per_day)
    return targets


def generate_schedule(
    start_date: date,
    weeks: int,
    ha: float,
    blocks_per_ha: int,
    trees_per_block: int,
    events_per_week: int,
    max_blocks_per_day: int,
    rest_on_sunday: bool = True,
    liters_per_tree_per_week: float = 15.0,
    water_split_mode: str = "events",  # "events" or "workdays"
):
    # total blocks
    total_blocks = int(round(ha * blocks_per_ha))
    if total_blocks <= 0:
        return pd.DataFrame(), {"total_blocks": 0, "warnings": ["ha または blocks_per_ha が0以下のため生成できません。"]}

    end_date = start_date + timedelta(days=weeks * 7 - 1)

    block_ids = make_block_ids(total_blocks)
    warnings = []
    rows = []

    for w in range(weeks):
        wk_start = start_date + timedelta(days=w * 7)
        wk_end = wk_start + timedelta(days=6)

        # this week's days (within horizon)
        week_days = [d for d in daterange(wk_start, wk_end) if start_date <= d <= end_date]
        if rest_on_sunday:
            week_days = [d for d in week_days if not is_sunday(d)]

        if not week_days:
            continue

        D = len(week_days)

        # required total assignments this week
        required = total_blocks * events_per_week
        capacity = D * max_blocks_per_day
        assignable = min(required, capacity)

        # targets per day for leveling
        day_targets = _even_targets(week_days, assignable, max_blocks_per_day)
        day_remaining = dict(day_targets)

        if required > capacity:
            warnings.append(
                f"Week {w+1}: 必要割当({required}件=ブロック{total_blocks}×{events_per_week}回/週)が、"
                f"週の上限枠({capacity}件=稼働日{D}×{max_blocks_per_day}ブロック/日)を超えています。"
                f" → max_blocks_per_day を増やすか、events_per_week を下げてください。"
            )

        # assignments per day
        assignments = {d: [] for d in week_days}
        # block -> set(days already assigned) (to avoid same-day duplicates if possible)
        block_day_used = {b: set() for b in block_ids}

        assigned_count = 0

        for k in range(events_per_week):
            # rotate blocks & shift day order to spread
            rotated_blocks = block_ids[k:] + block_ids[:k]
            day_order = week_days[(k % D):] + week_days[:(k % D)]
            day_idx = 0

            for b in rotated_blocks:
                if assigned_count >= assignable:
                    break

                # find a day with remaining capacity
                chosen_day = None

                # First pass: try to pick a day not used for this block yet
                tries = 0
                while tries < D:
                    d = day_order[day_idx % D]
                    if day_remaining.get(d, 0) > 0 and (d not in block_day_used[b]):
                        chosen_day = d
                        break
                    day_idx += 1
                    tries += 1

                # Second pass: if all days would duplicate for this block, allow duplicate but respect remaining capacity
                if chosen_day is None:
                    tries = 0
                    while tries < D:
                        d = day_order[day_idx % D]
                        if day_remaining.get(d, 0) > 0:
                            chosen_day = d
                            break
                        day_idx += 1
                        tries += 1

                if chosen_day is None:
                    # no capacity left
                    break

                assignments[chosen_day].append(b)
                block_day_used[b].add(chosen_day)
                day_remaining[chosen_day] -= 1
                assigned_count += 1
                day_idx += 1

        # per-tree liters for this schedule unit
        if water_split_mode == "workdays":
            liters_per_tree_per_unit = liters_per_tree_per_week / D
            water_unit_label = "稼働日割(ℓ/日)"
        else:
            liters_per_tree_per_unit = liters_per_tree_per_week / max(events_per_week, 1)
            water_unit_label = "潅水回数割(ℓ/回)"

        # rows per day
        for d in week_days:
            b_list = assignments.get(d, [])
            trees_est = len(b_list) * trees_per_block
            liters_total = trees_est * liters_per_tree_per_unit

            rows.append(
                {
                    "日付": d.isoformat(),
                    "曜日": WEEKDAYS_JA[d.weekday()],
                    "週": w + 1,
                    "ブロック数": len(b_list),
                    "ブロック一覧": ", ".join([str(x) for x in b_list]),
                    "本数(推定)": trees_est,
                    "配分方式": water_unit_label,
                    "1本あたり配分(ℓ)": round(liters_per_tree_per_unit, 2),
                    "必要水量(ℓ)": round(liters_total, 0),
                    "必要水量(m3)": round(liters_total / 1000.0, 2),
                }
            )

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
        "liters_per_tree_per_week": liters_per_tree_per_week,
        "water_split_mode": water_split_mode,
        "warnings": warnings,
    }
    return df, meta


def to_excel_bytes(df: pd.DataFrame, meta: dict):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Schedule")

        summary = pd.DataFrame(
            [
                ["ha", meta.get("ha")],
                ["blocks_per_ha", meta.get("blocks_per_ha")],
                ["trees_per_block", meta.get("trees_per_block")],
                ["total_blocks", meta.get("total_blocks")],
                ["events_per_week", meta.get("events_per_week")],
                ["max_blocks_per_day", meta.get("max_blocks_per_day")],
                ["weeks", meta.get("weeks")],
                ["rest_on_sunday", meta.get("rest_on_sunday")],
                ["liters_per_tree_per_week", meta.get("liters_per_tree_per_week")],
                ["water_split_mode", meta.get("water_split_mode")],
            ],
            columns=["項目", "値"],
        )
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
st.title("水やりスケジュール自動生成（平準化 / ブロック単位 / 日曜休み対応）")

with st.sidebar:
    st.header("入力")
    ha = st.number_input("対象面積（ha）", min_value=0.0, value=1.0, step=0.5)
    blocks_per_ha = st.number_input("1haあたりブロック数", min_value=1, value=17, step=1)
    trees_per_block = st.number_input("1ブロックあたり本数", min_value=1, value=117, step=1)
    start_date = st.date_input("開始日", value=date.today())
    weeks = st.number_input("何週間先まで作る？", min_value=1, value=8, step=1)
    events_per_week = st.selectbox("各ブロックを週に何回水やりする？", options=[1, 2, 3], index=0)
    max_blocks_per_day = st.number_input("1日に水やりする最大ブロック数（上限）", min_value=1, value=10, step=1)
    rest_on_sunday = st.checkbox("日曜は休み（固定）", value=True)

    st.divider()
    st.subheader("水量設定")
    liters_per_tree_per_week = st.number_input("1本あたり必要水量（ℓ/週）", min_value=0.0, value=15.0, step=1.0)
    water_split_mode = st.radio(
        "水量の配分方式（列の計算基準）",
        options=["潅水回数割（おすすめ）", "稼働日割（参考）"],
        index=0,
    )
    water_split_mode_key = "events" if water_split_mode.startswith("潅水回数割") else "workdays"

df, meta = generate_schedule(
    start_date=start_date,
    weeks=int(weeks),
    ha=float(ha),
    blocks_per_ha=int(blocks_per_ha),
    trees_per_block=int(trees_per_block),
    events_per_week=int(events_per_week),
    max_blocks_per_day=int(max_blocks_per_day),
    rest_on_sunday=bool(rest_on_sunday),
    liters_per_tree_per_week=float(liters_per_tree_per_week),
    water_split_mode=water_split_mode_key,
)

col1, col2 = st.columns([2, 1])

with col2:
    st.subheader("サマリー")
    st.write(f"- 対象面積: **{meta.get('ha', 0)} ha**")
    st.write(f"- 総ブロック数: **{meta.get('total_blocks', 0)} blocks**")
    st.write(f"- 1ブロック本数: **{meta.get('trees_per_block', 0)} 本**")
    st.write(f"- 週あたり回数: **{meta.get('events_per_week', 0)} 回/週**")
    st.write(f"- 1日上限: **{meta.get('max_blocks_per_day', 0)} blocks/日**")
    st.write(f"- 水量基準: **{meta.get('liters_per_tree_per_week', 0)} ℓ/本/週**")
    st.write(f"- 配分方式: **{meta.get('water_split_mode', '')}**")

    if meta.get("warnings"):
        st.error("⚠ 容量不足の可能性があります")
        for w in meta["warnings"]:
            st.write(f"- {w}")
    else:
        st.success("OK：週内に平準化して割り振り可能です")

with col1:
    st.subheader("生成されたスケジュール（平準化 + 水量列）")
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
``
