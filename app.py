"""
Crop Threat & Input Dashboard
------------------------------
Top track: crop growth stage reference timeline (always shown).
Bottom track: ONE swappable board at a time — Weed / Pest / Disease /
Fertilizer — selected from the sidebar. Only one board renders at once
by design (showing all four together is hard to read).

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

STAGE_COLORS = [
    "#8ECAE6", "#219EBC", "#023047", "#FFB703", "#FB8500",
    "#A7C957", "#6A994E", "#BC4749", "#9D4EDD", "#264653",
]

PALETTE = [
    "#457B9D", "#E76F51", "#2A9D8F", "#E9C46A", "#6A994E",
    "#BC4749", "#9D4EDD", "#F4A261", "#264653", "#A7C957",
]

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
            sheets[name] = pd.DataFrame()  # sheet not present
    return sheets


def get_file():
    st.sidebar.subheader("Data source")
    uploaded = st.sidebar.file_uploader("Upload workbook (.xlsx)", type=["xlsx"])
    if uploaded is not None:
        return uploaded
    if os.path.exists(DEFAULT_PATH):
        return DEFAULT_PATH
    return None


# ----------------------------------------------------------------------
# Chart builders
# ----------------------------------------------------------------------

def build_stage_timeline(stage_df: pd.DataFrame, label_col: str) -> go.Figure:
    """Single-track horizontal reference timeline of crop growth stages."""
    fig = go.Figure()
    stage_df = stage_df.sort_values("start_day").reset_index(drop=True)

    for i, row in stage_df.iterrows():
        duration = row["end_day"] - row["start_day"]
        color = STAGE_COLORS[i % len(STAGE_COLORS)]
        label = str(row[label_col])
        fig.add_trace(go.Bar(
            x=[duration],
            y=["Crop Stage"],
            base=[row["start_day"]],
            orientation="h",
            marker=dict(color=color, line=dict(color="white", width=1)),
            text=label,
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(color="white", size=13),
            hovertemplate=f"<b>{label}</b><br>Day {row['start_day']}–{row['end_day']}<extra></extra>",
            showlegend=False,
        ))

    fig.update_layout(
        barmode="stack",
        height=140,
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis=dict(showgrid=True, title="Day after planting"),
        yaxis=dict(showticklabels=False),
        title="Crop Growth Stage (reference)",
        bargap=0.4,
    )
    return fig


def build_board_timeline(df: pd.DataFrame, row_col: str, label_col: str,
                          color_col: str, hover_fn, title: str) -> go.Figure:
    """
    Generic swappable-board timeline: one row per `row_col` value, one bar
    per record (a row can have several bars if it has multiple windows),
    colored by the categories found in `color_col`.
    """
    if df.empty:
        fig = go.Figure()
        fig.update_layout(height=120, title=f"{title} — no data for this crop")
        return fig

    row_order = (
        df.groupby(row_col)["start_day"].min().sort_values().index.tolist()
    )

    color_values = sorted(df[color_col].dropna().astype(str).unique().tolist())
    color_map = {v: PALETTE[i % len(PALETTE)] for i, v in enumerate(color_values)}

    fig = go.Figure()
    seen_legend = set()
    for _, row in df.iterrows():
        duration = row["end_day"] - row["start_day"]
        cat = str(row.get(color_col, ""))
        color = color_map.get(cat, PALETTE[-1])
        label = str(row[label_col])
        show_legend = cat not in seen_legend and len(color_values) > 1
        seen_legend.add(cat)

        fig.add_trace(go.Bar(
            x=[duration],
            y=[row[row_col]],
            base=[row["start_day"]],
            orientation="h",
            marker=dict(color=color, line=dict(color="white", width=1)),
            text=label,
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(color="white", size=12, family="Georgia, serif"),
            hovertemplate=hover_fn(row),
            name=cat if cat else "—",
            legendgroup=cat,
            showlegend=show_legend,
        ))

    n_rows = max(len(row_order), 1)
    fig.update_layout(
        barmode="stack",
        height=90 + n_rows * 45,
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis=dict(showgrid=True, title="Day after planting"),
        yaxis=dict(categoryorder="array", categoryarray=row_order,
                   autorange="reversed", title=""),
        title=title,
        bargap=0.3,
        legend_title_text=color_col,
    )
    return fig


# ----------------------------------------------------------------------
# Board configs — how to turn each board's sheets into the generic chart
# ----------------------------------------------------------------------

def weed_board(crop_id, sheets):
    weeds = sheets["crop_weeds"]
    her = sheets["weed_her"]
    df = weeds[weeds["crop_id"] == crop_id].copy()
    her_c = her[her["crop_id"] == crop_id]
    df = df.merge(
        her_c[["ws_id", "weed_id", "common_name", "hrac_code"]],
        on=["ws_id", "weed_id"], how="left",
    )

    def hover(row):
        return (
            f"<b><i>{row['weed_science']}</i></b><br>"
            f"{row['weed_name_en']} / {row['weed_name_th']}<br>"
            f"Common name: {row.get('common_name', '')}<br>"
            f"HRAC code: {row.get('hrac_code', '')}<br>"
            f"Spray stage: {row.get('weed_stage', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<extra></extra>"
        )

    fig = build_board_timeline(df, row_col="weed_science", label_col="weed_science",
                                color_col="type", hover_fn=hover,
                                title="Weed Control Windows")
    detail_cols = ["weed_stage", "weed_science", "weed_name_en", "weed_name_th",
                    "common_name", "hrac_code", "type", "start_day", "end_day"]
    return fig, df, detail_cols


def pest_board(crop_id, sheets):
    pest = sheets["crop_pest"]
    ins = sheets["pest_ins"]
    df = pest[pest["crop_id"] == crop_id].copy()
    ins_c = ins[ins["crop_id"] == crop_id]
    df = df.merge(
        ins_c[["pest_id", "common_name", "irac_code"]],
        on="pest_id", how="left",
    )

    def hover(row):
        return (
            f"<b>{row['pest_name_en']}</b><br>"
            f"{row['pest_name_th']}<br>"
            f"Order: {row.get('order', '')}<br>"
            f"Common name: {row.get('common_name', '')}<br>"
            f"IRAC code: {row.get('irac_code', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<extra></extra>"
        )

    fig = build_board_timeline(df, row_col="pest_name_en", label_col="pest_name_en",
                                color_col="order", hover_fn=hover,
                                title="Pest Pressure Windows")
    detail_cols = ["pest_name_en", "pest_name_th", "order", "common_name",
                   "irac_code", "start_day", "end_day"]
    return fig, df, detail_cols


def disease_board(crop_id, sheets):
    dis = sheets["crop_disease"]
    fun = sheets["disease_fun"]
    df = dis[dis["crop_id"] == crop_id].copy()
    fun_c = fun[fun["crop_id"] == crop_id]
    df = df.merge(
        fun_c[["disease_id", "common_name", "frac_code"]],
        on="disease_id", how="left",
    )

    def hover(row):
        return (
            f"<b><i>{row['disease_name_sc']}</i></b><br>"
            f"{row['disease_name_en']} / {row['disease_name_th']}<br>"
            f"Common name: {row.get('common_name', '')}<br>"
            f"FRAC code: {row.get('frac_code', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<extra></extra>"
        )

    fig = build_board_timeline(df, row_col="disease_name_sc", label_col="disease_name_sc",
                                color_col="type", hover_fn=hover,
                                title="Disease Pressure Windows")
    detail_cols = ["disease_name_sc", "disease_name_en", "disease_name_th",
                   "common_name", "frac_code", "type", "start_day", "end_day"]
    return fig, df, detail_cols


def fertilizer_board(crop_id, sheets):
    fert = sheets["fertilizer"]
    df = fert[fert["crop_id"] == crop_id].copy()
    df["_none"] = "Fertilizer"  # single-category color column

    def hover(row):
        return (
            f"<b>{row['formula']}</b><br>"
            f"Day {row['start_day']}–{row['end_day']}<extra></extra>"
        )

    fig = build_board_timeline(df, row_col="formula", label_col="formula",
                                color_col="_none", hover_fn=hover,
                                title="Fertilizer Application Windows")
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

st.plotly_chart(build_stage_timeline(crop_stage_df, label_col), use_container_width=True)

fig, board_df, detail_cols = BOARDS[board_choice](crop_id, sheets)
st.plotly_chart(fig, use_container_width=True)

if board_df.empty:
    st.info(f"No {board_choice.lower()} data for this crop.")
else:
    with st.expander(f"{board_choice} detail table"):
        st.dataframe(board_df[detail_cols], use_container_width=True, hide_index=True)