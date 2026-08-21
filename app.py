"""
Crop Threat & Input Dashboard
------------------------------
Top track: a continuous rice growth timeline (0 -> last day), stage
boundaries marked as ruler ticks, stage names as labels between ticks —
like a ruler, not separate colored chips.

Bottom track: ONE swappable board at a time — Weed / Pest / Disease /
Fertilizer — selected from the sidebar. Each real-world window (a weed's
pre-emergence window, a pest's pressure window, etc.) is ONE box, even if
several chemicals/products apply to it — those are combined into that
single box's label/hover, not drawn as separate side-by-side boxes.
Only genuinely different time windows (e.g. two separate spray dates for
the same weed) get their own box / sub-lane.

Expected workbook: crop_timeline.xlsx, with sheets:
  crop_stage    : crop_id, crop, stage, stage_th, start_day, end_day
  crop_weeds    : crop_id, ws_id, weed_stage, weed_id, weed_name_en,
                  weed_name_th, weed_science, type, start_day, end_day
  weed_her      : crop_id, ws_id, weed_id, weed_name_en, weed_name_th,
                  common_name, hrac_code
  crop_pest     : crop_id, pest_id, pest_name_en, pest_name_th, order,
                  start_day, end_day
  pest_ins      : crop_id, pest_id, pest_name_th, common_name, irac_code
  crop_disease  : crop_id, disease_id, disease_name_en, disease_name_th,
                  disease_name_sc, type, start_day, end_day
  disease_fun   : crop_id, disease_id, disease_name_th, common_name,
                  frac_code
  fertilizer    : crop_id, crop, formula, start_day, end_day
"""

import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Crop Threat & Input Dashboard", layout="wide")

DEFAULT_PATH = "crop_timeline.xlsx"

SHEET_NAMES = [
    "crop_stage", "crop_weeds", "weed_her",
    "crop_pest", "pest_ins",
    "crop_disease", "disease_fun",
    "fertilizer",
]

PALETTE = [
    "#457B9D", "#E76F51", "#2A9D8F", "#E9C46A", "#6A994E",
    "#BC4749", "#9D4EDD", "#F4A261", "#264653", "#A7C957",
]

RULER_COLOR = "#CFE8D5"
RULER_LINE = "#2E7D32"
RULER_TEXT = "#1B4332"

NARROW_DAY_THRESHOLD = 6  # below this many days, rotate the label vertical

# ----------------------------------------------------------------------
# Data loading
# ----------------------------------------------------------------------

@st.cache_data
def load_workbook(file):
    sheets = {}
    for name in SHEET_NAMES:
        try:
            df = pd.read_excel(file, sheet_name=name)
            df.columns = [c.strip() for c in df.columns]
            sheets[name] = df
        except ValueError:
            sheets[name] = pd.DataFrame()
    return sheets


def get_file():
    st.sidebar.subheader("Data source")
    uploaded = st.sidebar.file_uploader("Upload workbook (.xlsx)", type=["xlsx"])
    if st.sidebar.button("🔄 Reload data"):
        st.cache_data.clear()
        st.rerun()
    if uploaded is not None:
        return uploaded
    if os.path.exists(DEFAULT_PATH):
        return DEFAULT_PATH
    return None


# ----------------------------------------------------------------------
# Lane assignment — only kicks in for genuinely different, overlapping
# time windows within the same row (e.g. two distinct spray dates for
# the same weed). Duplicate chemical rows for the SAME window are merged
# upstream before this ever runs.
# ----------------------------------------------------------------------

def assign_lanes(group: pd.DataFrame):
    lanes_end = []
    assignment = {}
    for idx, row in group.sort_values("start_day").iterrows():
        placed = False
        for lane_idx in range(len(lanes_end)):
            if row["start_day"] >= lanes_end[lane_idx]:
                lanes_end[lane_idx] = row["end_day"]
                assignment[idx] = lane_idx
                placed = True
                break
        if not placed:
            lanes_end.append(row["end_day"])
            assignment[idx] = len(lanes_end) - 1
    return assignment, max(len(lanes_end), 1)


# ----------------------------------------------------------------------
# Chemical aggregation — collapse multiple chemical/product rows that
# belong to the SAME window (same weed/pest/disease + same start/end)
# into one row with a combined chemical list, so the chart draws ONE box.
# ----------------------------------------------------------------------

def aggregate_chemicals(merged: pd.DataFrame, group_cols: list,
                         name_col: str, code_col: str, code_label: str):
    def _agg(g):
        pairs = [
            (str(n).strip(), str(c).strip())
            for n, c in zip(g[name_col], g[code_col])
            if pd.notna(n) or pd.notna(c)
        ]
        pairs = [p for p in pairs if p[0] not in ("", "nan") or p[1] not in ("", "nan")]
        count = len(pairs)
        if count:
            chem_html = "<br>".join(f"• {n} ({code_label} {c})" for n, c in pairs)
        else:
            chem_html = "—"
        return pd.Series({"chem_count": count, "chem_list_html": chem_html})

    agg = merged.groupby(group_cols, dropna=False).apply(_agg).reset_index()
    return agg


# ----------------------------------------------------------------------
# Generic chart engine for the swappable boards
# ----------------------------------------------------------------------

def build_timeline_chart(df: pd.DataFrame, row_col: str, label_col: str,
                          color_col: str, hover_fn, title: str,
                          stage_df: pd.DataFrame = None, stage_label_col: str = "stage",
                          show_legend: bool = True) -> go.Figure:
    """
    Single plot. The crop growth stage isn't drawn as separate chart
    elements — it's just the x-axis tick labels: each stage's name plus
    its day range, positioned at that stage's midpoint. The board's
    rows/boxes are drawn against that labeled axis.
    """
    if df.empty:
        fig = go.Figure()
        fig.update_layout(height=120, title=f"{title} — no data for this crop")
        return fig

    row_order = df.groupby(row_col)["start_day"].min().sort_values().index.tolist()
    row_to_base = {r: i for i, r in enumerate(row_order)}
    n_rows = len(row_order)

    color_values = sorted(df[color_col].dropna().astype(str).unique().tolist())
    color_map = {v: PALETTE[i % len(PALETTE)] for i, v in enumerate(color_values)}
    multi_category = len(color_values) > 1

    fig = go.Figure()

    # --- board rows/boxes ---
    seen_legend = set()
    row_lane_counts = {}
    for row_val, group in df.groupby(row_col):
        lane_map, n_lanes = assign_lanes(group)
        row_lane_counts[row_val] = n_lanes
        base_y = row_to_base[row_val]
        lane_height = min(0.8 / n_lanes, 0.5)

        for idx, lane in lane_map.items():
            row = df.loc[idx]
            duration = row["end_day"] - row["start_day"]
            y_center = base_y + (lane - (n_lanes - 1) / 2) * lane_height
            cat = str(row.get(color_col, ""))
            color = color_map.get(cat, PALETTE[-1])
            show_this_legend = multi_category and cat not in seen_legend
            seen_legend.add(cat)

            fig.add_trace(go.Bar(
                x=[duration],
                y=[y_center],
                base=[row["start_day"]],
                orientation="h",
                width=lane_height * 0.85,
                marker=dict(color=color, line=dict(color="white", width=1)),
                hovertemplate=hover_fn(row),
                name=cat if cat else "—",
                legendgroup=cat,
                showlegend=show_this_legend,
            ))

    total_lane_rows = sum(row_lane_counts.values())

    # --- x-axis: regular day ticks (0, 20, 40, ...) as the main axis,
    #     with stage names as a second label row underneath. (No axis
    #     title text here — it was sitting in the same spot as the
    #     stage-name row and hiding it.) ---
    xaxis = dict(showgrid=True)
    stage_annotations = []
    if stage_df is not None and not stage_df.empty:
        sdf = stage_df.sort_values("start_day").reset_index(drop=True)
        stage_min = float(sdf["start_day"].min())
        stage_max = float(sdf["end_day"].max())
        span = stage_max - stage_min

        step = 20
        day_ticks = list(range(0, int(stage_max) + 1, step))
        if not day_ticks or day_ticks[-1] != int(stage_max):
            day_ticks.append(int(stage_max))

        xaxis.update(
            tickmode="array",
            tickvals=day_ticks,
            ticktext=[str(t) for t in day_ticks],
            range=[stage_min - span * 0.02, stage_max + span * 0.02],
        )

        for _, r in sdf.iterrows():
            mid = (r["start_day"] + r["end_day"]) / 2
            stage_annotations.append(dict(
                x=mid, y=-0.14, xref="x", yref="paper",
                text=str(r[stage_label_col]), showarrow=False,
                font=dict(color=RULER_TEXT, size=12, family="Georgia, serif"),
                yanchor="top",
            ))
        stage_annotations.append(dict(
            x=0, y=-0.14, xref="paper", yref="paper",
            text="Day after planting / Stage:", showarrow=False,
            xanchor="left", yanchor="top",
            font=dict(color="#666666", size=10),
        ))

    fig.update_layout(
        barmode="overlay",
        height=max(220, 130 + total_lane_rows * 42),
        margin=dict(l=10, r=10, t=45, b=55),
        xaxis=xaxis,
        yaxis=dict(
            tickmode="array",
            tickvals=[row_to_base[r] for r in row_order],
            ticktext=row_order,
            range=[n_rows - 0.5, -0.5],
            title="",
        ),
        title=title,
        annotations=stage_annotations,
        showlegend=multi_category and show_legend,
        legend_title_text=color_col,
    )
    return fig


# ----------------------------------------------------------------------
# Board configs
# ----------------------------------------------------------------------

def weed_board(crop_id, sheets, crop_stage_df, stage_label_col):
    weeds = sheets["crop_weeds"]
    her = sheets["weed_her"]
    raw = weeds[weeds["crop_id"] == crop_id].copy()
    her_c = her[her["crop_id"] == crop_id]
    merged = raw.merge(
        her_c[["ws_id", "weed_id", "common_name", "hrac_code"]],
        on=["ws_id", "weed_id"], how="left",
    )

    group_cols = ["crop_id", "ws_id", "weed_id", "weed_stage", "weed_science",
                  "weed_name_en", "weed_name_th", "type", "start_day", "end_day"]
    agg = aggregate_chemicals(merged, group_cols, "common_name", "hrac_code", "HRAC")
    df = merged[group_cols].drop_duplicates().merge(agg, on=group_cols)

    def hover(row):
        return (
            f"<b><i>{row['weed_science']}</i></b><br>"
            f"{row['weed_name_en']} / {row['weed_name_th']}<br>"
            f"Stage: {row.get('weed_stage', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br>"
            f"<br><b>Products:</b><br>{row['chem_list_html']}"
            "<extra></extra>"
        )

    fig = build_timeline_chart(df, row_col="weed_science", label_col="weed_stage",
                                color_col="type", hover_fn=hover,
                                title="Weed Control Windows",
                                stage_df=crop_stage_df, stage_label_col=stage_label_col)
    detail_cols = ["weed_stage", "weed_science", "weed_name_en", "weed_name_th",
                    "common_name", "hrac_code", "type", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def pest_board(crop_id, sheets, crop_stage_df, stage_label_col):
    pest = sheets["crop_pest"]
    ins = sheets["pest_ins"]
    raw = pest[pest["crop_id"] == crop_id].copy()
    ins_c = ins[ins["crop_id"] == crop_id]
    merged = raw.merge(
        ins_c[["pest_id", "common_name", "irac_code"]],
        on="pest_id", how="left",
    )

    group_cols = ["crop_id", "pest_id", "pest_name_en", "pest_name_th",
                  "order", "start_day", "end_day"]
    agg = aggregate_chemicals(merged, group_cols, "common_name", "irac_code", "IRAC")
    df = merged[group_cols].drop_duplicates().merge(agg, on=group_cols)

    def hover(row):
        return (
            f"<b>{row['pest_name_en']}</b><br>"
            f"{row['pest_name_th']}<br>"
            f"Order: {row.get('order', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br>"
            f"<br><b>Products:</b><br>{row['chem_list_html']}"
            "<extra></extra>"
        )

    fig = build_timeline_chart(df, row_col="pest_name_en", label_col="pest_name_en",
                                color_col="order", hover_fn=hover,
                                title="Pest Pressure Windows",
                                stage_df=crop_stage_df, stage_label_col=stage_label_col)
    detail_cols = ["pest_name_en", "pest_name_th", "order", "common_name",
                   "irac_code", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def disease_board(crop_id, sheets, crop_stage_df, stage_label_col):
    dis = sheets["crop_disease"]
    fun = sheets["disease_fun"]
    raw = dis[dis["crop_id"] == crop_id].copy()
    fun_c = fun[fun["crop_id"] == crop_id]
    merged = raw.merge(
        fun_c[["disease_id", "common_name", "frac_code"]],
        on="disease_id", how="left",
    )

    group_cols = ["crop_id", "disease_id", "disease_name_en", "disease_name_th",
                  "disease_name_sc", "type", "start_day", "end_day"]
    agg = aggregate_chemicals(merged, group_cols, "common_name", "frac_code", "FRAC")
    df = merged[group_cols].drop_duplicates().merge(agg, on=group_cols)

    def hover(row):
        return (
            f"<b><i>{row['disease_name_sc']}</i></b><br>"
            f"{row['disease_name_en']} / {row['disease_name_th']}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br>"
            f"<br><b>Products:</b><br>{row['chem_list_html']}"
            "<extra></extra>"
        )

    fig = build_timeline_chart(df, row_col="disease_name_sc", label_col="disease_name_sc",
                                color_col="type", hover_fn=hover,
                                title="Disease Pressure Windows",
                                stage_df=crop_stage_df, stage_label_col=stage_label_col)
    detail_cols = ["disease_name_sc", "disease_name_en", "disease_name_th",
                   "common_name", "frac_code", "type", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def fertilizer_board(crop_id, sheets, crop_stage_df, stage_label_col):
    fert = sheets["fertilizer"]
    df = fert[fert["crop_id"] == crop_id].copy()
    df["_none"] = "Fertilizer"

    def hover(row):
        return (
            f"<b>{row['formula']}</b><br>"
            f"Day {row['start_day']}–{row['end_day']}<extra></extra>"
        )

    fig = build_timeline_chart(df, row_col="formula", label_col="formula",
                                color_col="_none", hover_fn=hover,
                                title="Fertilizer Application Windows",
                                stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                show_legend=False)
    detail_cols = ["formula", "start_day", "end_day"]
    return fig, df, detail_cols


BOARDS = {
    "Weed": weed_board,
    "Pest": pest_board,
    "Disease": disease_board,
    "Fertilizer": fertilizer_board,
}

# ----------------------------------------------------------------------
# App
# ----------------------------------------------------------------------

st.title("🌾 Crop Threat & Input Dashboard")

board_choice = st.sidebar.radio("Board", list(BOARDS.keys()), index=0)

data_file = get_file()
if data_file is None:
    st.warning(
        f"No workbook found. Upload one from the sidebar, or place a file "
        f"named `{DEFAULT_PATH}` next to `app.py`."
    )
    st.stop()

try:
    sheets = load_workbook(data_file)
except Exception as e:
    st.error(f"Couldn't read the workbook: {e}")
    st.stop()

stage_df_all = sheets["crop_stage"]
if stage_df_all.empty:
    st.error("`crop_stage` sheet is missing or empty.")
    st.stop()

crop_lookup = stage_df_all[["crop_id", "crop"]].drop_duplicates()
crop_name_to_id = dict(zip(crop_lookup["crop"], crop_lookup["crop_id"]))

col1, col2 = st.columns([2, 1])
with col1:
    crop_choice = st.selectbox("Crop", list(crop_name_to_id.keys()))
with col2:
    stage_label_choice = st.radio("Stage label language", ["English", "Thai"], horizontal=True)
label_col = "stage" if stage_label_choice == "English" else "stage_th"

crop_id = crop_name_to_id[crop_choice]

crop_stage_df = stage_df_all[stage_df_all["crop_id"] == crop_id]
if crop_stage_df.empty:
    st.warning("No stage data for this crop.")
    st.stop()

fig, board_df, detail_cols = BOARDS[board_choice](crop_id, sheets, crop_stage_df, label_col)
st.plotly_chart(fig, use_container_width=True)

if board_df.empty:
    st.info(f"No {board_choice.lower()} data for this crop.")
else:
    with st.expander(f"{board_choice} detail table (one row per product)"):
        st.dataframe(board_df, use_container_width=True, hide_index=True)
