"""
Crop Threat & Input Dashboard
------------------------------
Top track: a continuous rice growth timeline (0 -> last day), stage
boundaries marked as ruler ticks, stage names as labels between ticks —
like a ruler, not separate colored chips.

Bottom track: ONE swappable board at a time — Weed / Insect / Disease /
Fertilizer — selected from the sidebar. Each real-world window (a weed's
pre-emergence window, an insect's pressure window, etc.) is ONE box, even if
several chemicals/products apply to it — those are combined into that
single box's label/hover, not drawn as separate side-by-side boxes.
Only genuinely different time windows (e.g. two separate spray dates for
the same weed) get their own box / sub-lane.

The English/Thai toggle in the sidebar now drives BOTH the crop-stage
row labels AND the weed/insect/disease row labels on the y-axis. Fertilizer
row labels (the formula name) are language-agnostic and never change.

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

username = st.text_input("Username")
password = st.text_input("Password", type="password")

if st.button("Login"):
    if (
        username == st.secrets["login"]["username"]
        and password == st.secrets["login"]["password"]
    ):
        st.session_state["logged_in"] = True

if st.session_state.get("logged_in"):
    st.write("Welcome")

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

STAGE_COLORS = [
    "#8ECAE6", "#219EBC", "#023047", "#FFB703", "#FB8500",
    "#A7C957", "#6A994E", "#BC4749", "#9D4EDD", "#264653",
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
            # Drop the stray "Unnamed: N" spillover columns from blank
            # trailing cells in the source workbook — not part of the
            # real schema, just noise from Excel's used-range detection.
            df = df.loc[:, ~df.columns.str.startswith("Unnamed:")]
            # Trim stray leading/trailing whitespace from every text cell —
            # otherwise "granular" and "granular " (or similar) are treated
            # as two different categories (duplicate dropdown entries,
            # duplicate legend colors, broken row grouping, etc.).
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
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
# belong to the SAME window (same weed/insect/disease + same start/end)
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
            chem_html = "<br>".join(f"• {n} ({c})" for n, c in pairs)
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
                          show_legend: bool = True, row_label_map: dict = None) -> go.Figure:
    """
    Single plot. Crop growth stage is drawn as a real row at the top of
    the SAME chart (same x/y coordinate space as the weed/insect/disease
    boxes) — not a floating annotation outside the plot area, which was
    getting clipped by the renderer.

    row_col is the (language-invariant) grouping/identity key for each
    row — e.g. a weed's scientific name, an insect's English name, a
    disease's scientific name. row_label_map, if given, maps each of
    those row_col values to the display label the user should actually
    see on the y-axis (e.g. the Thai common name when the Thai toggle is
    selected). If row_label_map is omitted, row_col's own values are
    used as the display label (this is how the Fertilizer board — whose
    row identity, the formula name, doesn't change with language — keeps
    working unmodified).
    """
    if df.empty:
        fig = go.Figure()
        fig.update_layout(height=120, title=f"{title} — no data for this crop")
        return fig

    order_df = (
        df.groupby(row_col)
        .agg(**{color_col: (color_col, "first"), "start_day": ("start_day", "min")})
        .reset_index()
        .sort_values([color_col, "start_day"])
    )
    row_order = order_df[row_col].tolist()
    row_to_base = {r: i for i, r in enumerate(row_order)}
    n_rows = len(row_order)

    color_values = sorted(df[color_col].dropna().astype(str).unique().tolist())
    color_map = {v: PALETTE[i % len(PALETTE)] for i, v in enumerate(color_values)}
    multi_category = len(color_values) > 1

    fig = go.Figure()
    annotations = []

    # --- crop stage: a real row at the top (y = -1.3), drawn with the
    #     same bar/annotation mechanism as everything else below it ---
    STAGE_ROW_Y = -1.3
    top_of_axis = -0.5
    if stage_df is not None and not stage_df.empty:
        sdf = stage_df.sort_values("start_day").reset_index(drop=True)
        for i, srow in sdf.iterrows():
            duration = srow["end_day"] - srow["start_day"]
            fig.add_trace(go.Bar(
                x=[duration], y=[STAGE_ROW_Y], base=[srow["start_day"]],
                orientation="h", width=0.7,
                marker=dict(color=STAGE_COLORS[i % len(STAGE_COLORS)],
                            line=dict(color="white", width=1)),
                hovertemplate=f"<b>{srow[stage_label_col]}</b><br>Day "
                               f"{srow['start_day']}–{srow['end_day']}<extra></extra>",
                showlegend=False,
            ))
            mid = (srow["start_day"] + srow["end_day"]) / 2
            annotations.append(dict(
                x=mid, y=STAGE_ROW_Y, xref="x", yref="y",
                text=str(srow[stage_label_col]), showarrow=False,
                font=dict(color="white", size=17, family="Georgia, serif"),
                xanchor="center", yanchor="middle",
            ))
        top_of_axis = STAGE_ROW_Y - 0.8

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

    # --- x-axis: plain day-number ticks (0, 20, 40, ...) ---
    xaxis = dict(showgrid=True, title=dict(text="Day after planting", font=dict(size=19)),
                 tickfont=dict(size=18))
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

    y_ticks = [row_to_base[r] for r in row_order]
    if row_label_map:
        y_ticktext = [row_label_map.get(r, r) for r in row_order]
    else:
        y_ticktext = list(row_order)
    if stage_df is not None and not stage_df.empty:
        y_ticks = [STAGE_ROW_Y] + y_ticks
        y_ticktext = ["Crop Stage"] + y_ticktext

    fig.update_layout(
        barmode="overlay",
        height=max(240, 150 + total_lane_rows * 54),
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis=xaxis,
        yaxis=dict(
            tickmode="array",
            tickvals=y_ticks,
            ticktext=y_ticktext,
            range=[n_rows - 0.5, top_of_axis],
            title="",
            tickfont=dict(size=19),
            automargin=True,
        ),
        annotations=annotations,
        showlegend=multi_category and show_legend,
        legend_title_text=color_col,
        legend=dict(font=dict(size=17)),
        hoverlabel=dict(font=dict(size=20), align="left"),
        font=dict(size=17),
    )
    return fig


# ----------------------------------------------------------------------
# Board configs
# ----------------------------------------------------------------------

def weed_board(crop_id, sheets, crop_stage_df, stage_label_col):
    is_thai = stage_label_col.endswith("_th")
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

    # Row identity stays on the language-invariant scientific name.
    # English display label uses weed_science itself (not weed_name_en);
    # Thai display label still switches to weed_name_th.
    name_col = "weed_name_th" if is_thai else "weed_science"
    row_label_map = dict(zip(df["weed_science"], df[name_col]))

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
                                stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                row_label_map=row_label_map)
    detail_cols = ["weed_stage", "weed_science", "weed_name_en", "weed_name_th",
                    "common_name", "hrac_code", "type", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def insect_board(crop_id, sheets, crop_stage_df, stage_label_col):
    is_thai = stage_label_col.endswith("_th")
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

    # Row identity stays on the English name (language-invariant key);
    # only the displayed y-axis label switches with the toggle.
    name_col = "pest_name_th" if is_thai else "pest_name_en"
    row_label_map = dict(zip(df["pest_name_en"], df[name_col]))

    def hover(row):
        return (
            f"<b>{row['pest_name_en']}</b><br>"
            f"{row['pest_name_th']}<br>"
            f"Insect order: {row.get('order', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br>"
            f"<br><b>Products:</b><br>{row['chem_list_html']}"
            "<extra></extra>"
        )

    fig = build_timeline_chart(df, row_col="pest_name_en", label_col="pest_name_en",
                                color_col="order", hover_fn=hover,
                                title="Insect Pressure Windows",
                                stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                row_label_map=row_label_map)
    detail_cols = ["pest_name_en", "pest_name_th", "order", "common_name",
                   "irac_code", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def disease_board(crop_id, sheets, crop_stage_df, stage_label_col):
    is_thai = stage_label_col.endswith("_th")
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

    # Row identity stays on the language-invariant scientific name;
    # only the displayed y-axis label switches with the toggle.
    name_col = "disease_name_th" if is_thai else "disease_name_en"
    row_label_map = dict(zip(df["disease_name_sc"], df[name_col]))

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
                                stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                row_label_map=row_label_map)
    detail_cols = ["disease_name_sc", "disease_name_en", "disease_name_th",
                   "common_name", "frac_code", "type", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def fertilizer_board(crop_id, sheets, crop_stage_df, stage_label_col):
    # Backward/forward compatible: works with the old sheet (just
    # formula/start_day/end_day) as well as the new one that adds
    # `stage` (e.g. "First application", "Second application") and
    # `type` (e.g. "granular", "foliar").
    is_thai = stage_label_col.endswith("_th")
    fert = sheets["fertilizer"]
    df = fert[fert["crop_id"] == crop_id].copy()

    has_stage = "stage" in df.columns
    has_type = "type" in df.columns and df["type"].notna().any()

    # --- type filter + "total use" summary (e.g. "foliar + granular") ---
    # Computed on the un-aggregated, one-row-per-formula data.
    if has_type:
        type_options = sorted(df["type"].dropna().astype(str).unique().tolist())
        selected_types = st.multiselect(
            "Fertilizer type", type_options, default=type_options,
            help="Choose one type, or keep several selected to see them combined "
                 "on the same timeline (e.g. foliar + granular).",
        )
        df = df[df["type"].astype(str).isin(selected_types)] if selected_types else df.iloc[0:0]

    detail_cols = [c for c in ["stage", "type", "formula", "start_day", "end_day"]
                   if c in df.columns]
    detail_df = df[detail_cols].copy()

    if df.empty:
        fig = build_timeline_chart(df, row_col="formula", label_col="formula",
                                    color_col="_none", hover_fn=lambda r: "",
                                    title="Fertilizer Application Windows",
                                    stage_df=crop_stage_df, stage_label_col=stage_label_col)
        return fig, detail_df, detail_cols

    # --- row identity: stage name on the y-axis ---
    row_col = "stage" if has_stage else "formula"
    if has_stage:
        stage_name_col = "stage_th" if (is_thai and "stage_th" in df.columns) else "stage"
        row_label_map = dict(zip(df[row_col], df[stage_name_col]))
    else:
        row_label_map = None  # falls back to showing the formula itself

    # --- collapse every formula that shares the SAME window (same stage,
    #     same start/end day) into ONE box with a combined formula list,
    #     the same way Weed/Insect/Disease combine multiple products. ---
    group_cols = [c for c in ["crop_id", row_col, "start_day", "end_day"] if c in df.columns]

    def _agg(g):
        if has_type:
            items = [
                f"• {f} ({t})" for f, t in zip(g["formula"], g["type"])
                if pd.notna(f) or pd.notna(t)
            ]
            types_present = sorted({str(t) for t in g["type"].dropna()})
        else:
            items = [f"• {f}" for f in g["formula"] if pd.notna(f)]
            types_present = []
        return pd.Series({
            "formula_list_html": "<br>".join(items) if items else "—",
            "type_combo": " + ".join(types_present) if types_present else "Fertilizer",
        })

    agg = df.groupby(group_cols, dropna=False).apply(_agg).reset_index()
    df_agg = df[group_cols].drop_duplicates().merge(agg, on=group_cols)

    color_col = "type_combo"

    def hover(row):
        parts = []
        if has_stage:
            parts.append(f"<b>{row['stage']}</b>")
        parts.append(f"Day {row['start_day']}–{row['end_day']}")
        parts.append(f"<br><b>Formula:</b><br>{row['formula_list_html']}")
        return "<br>".join(parts) + "<extra></extra>"

    fig = build_timeline_chart(df_agg, row_col=row_col, label_col=row_col,
                                color_col=color_col, hover_fn=hover,
                                title="Fertilizer Application Windows",
                                stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                show_legend=has_type, row_label_map=row_label_map)
    return fig, detail_df, detail_cols


BOARDS = {
    "Weed": weed_board,
    "Insect": insect_board,
    "Disease": disease_board,
    "Fertilizer": fertilizer_board,
}

# ----------------------------------------------------------------------
# App
# ----------------------------------------------------------------------

st.title("🌾 Crop Threat & Input Dashboard")

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

col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    crop_choice = st.selectbox("Crop", list(crop_name_to_id.keys()))
with col2:
    board_choice = st.selectbox("Board", list(BOARDS.keys()), index=0)
with col3:
    stage_label_choice = st.radio("Label language", ["English", "Thai"], horizontal=True)
label_col = "stage" if stage_label_choice == "English" else "stage_th"

crop_id = crop_name_to_id[crop_choice]

crop_stage_df = stage_df_all[stage_df_all["crop_id"] == crop_id]
if crop_stage_df.empty:
    st.warning("No stage data for this crop.")
    st.stop()

BOARD_TITLES = {
    "Weed": "Weed Control Windows",
    "Insect": "Insect Pressure Windows",
    "Disease": "Disease Pressure Windows",
    "Fertilizer": "Fertilizer Application Windows",
}

st.subheader(BOARD_TITLES[board_choice])
fig, board_df, detail_cols = BOARDS[board_choice](crop_id, sheets, crop_stage_df, label_col)
st.plotly_chart(fig, use_container_width=True)

if board_df.empty:
    st.info(f"No {board_choice.lower()} data for this crop.")
else:
    with st.expander(f"{board_choice} detail table (one row per product)"):
        st.dataframe(board_df, use_container_width=True, hide_index=True)
