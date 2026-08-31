"""
SAC Crop Dashboard — combined app
----------------------------------
One app, one login, sidebar switch between two independent views that
each load their own workbook:

  "Crop Threat & Input" -> crop_timeline.xlsx
      Boxes colored by chemical/threat type. See THREAT section below
      for the expected sheet schema.

  "Product Coverage"    -> crop_timeline_coverage.xlsx
      Boxes colored by whether the selected company has a product for
      that window (green) or not (red). See COVERAGE section below for
      the expected sheet schema.

The two workbooks are completely independent — different files, loaded
and cached separately — so changes to one never affect the other.
"""

import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="SAC Crop Dashboard", layout="wide")

# =====================================================================
# Shared constants / helpers used by both views
# =====================================================================

STAGE_COLORS = [
    "#8ECAE6", "#219EBC", "#023047", "#FFB703", "#FB8500",
    "#A7C957", "#6A994E", "#BC4749", "#9D4EDD", "#264653",
]


def assign_lanes(group: pd.DataFrame):
    """For genuinely different, overlapping time windows within the same
    row (e.g. two distinct spray dates for the same weed)."""
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


# =====================================================================
# ============================ THREAT VIEW ============================
# Reads crop_timeline.xlsx:
#   crop_stage    : crop_id, crop, stage, stage_th, start_day, end_day
#   crop_weeds    : crop_id, ws_id, weed_stage, weed_id, weed_name_en,
#                   weed_name_th, weed_science, type, start_day, end_day
#   weed_her      : crop_id, ws_id, weed_id, weed_name_en, weed_name_th,
#                   common_name, hrac_code
#   crop_pest     : crop_id, pest_id, pest_name_en, pest_name_th, order,
#                   start_day, end_day
#   pest_ins      : crop_id, pest_id, pest_name_th, common_name, irac_code
#   crop_disease  : crop_id, disease_id, disease_name_en, disease_name_th,
#                   disease_name_sc, type, start_day, end_day
#   disease_fun   : crop_id, disease_id, disease_name_th, common_name,
#                   frac_code
#   fertilizer    : crop_id, crop, formula, start_day, end_day
# =====================================================================

DEFAULT_PATH_THREAT = "crop_timeline.xlsx"

SHEET_NAMES_THREAT = [
    "crop_stage", "crop_weeds", "weed_her",
    "crop_pest", "pest_ins",
    "crop_disease", "disease_fun",
    "fertilizer",
]

PALETTE = [
    "#457B9D", "#E76F51", "#2A9D8F", "#E9C46A", "#6A994E",
    "#BC4749", "#9D4EDD", "#F4A261", "#264653", "#A7C957",
]

BOARD_TITLES_THREAT = {
    "Weed": "Weed Control Windows",
    "Insect": "Insect Pressure Windows",
    "Disease": "Disease Pressure Windows",
    "Fertilizer": "Fertilizer Application Windows",
}


@st.cache_data
def load_workbook_threat(file):
    sheets = {}
    for name in SHEET_NAMES_THREAT:
        try:
            df = pd.read_excel(file, sheet_name=name)
            df.columns = [c.strip() for c in df.columns]
            df = df.loc[:, ~df.columns.str.startswith("Unnamed:")]
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
            sheets[name] = df
        except ValueError:
            sheets[name] = pd.DataFrame()
    return sheets


def get_file_threat():
    st.sidebar.subheader("Threat & Input data source")
    uploaded = st.sidebar.file_uploader(
        "Upload workbook (.xlsx)", type=["xlsx"], key="threat_uploader"
    )
    if st.sidebar.button("🔄 Reload data", key="threat_reload"):
        st.cache_data.clear()
        st.rerun()
    if uploaded is not None:
        return uploaded
    if os.path.exists(DEFAULT_PATH_THREAT):
        return DEFAULT_PATH_THREAT
    return None


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


def build_timeline_chart_threat(df: pd.DataFrame, row_col: str, label_col: str,
                                 color_col: str, hover_fn, title: str,
                                 stage_df: pd.DataFrame = None, stage_label_col: str = "stage",
                                 show_legend: bool = True, row_label_map: dict = None) -> go.Figure:
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


def weed_board_threat(crop_id, sheets, crop_stage_df, stage_label_col):
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

    fig = build_timeline_chart_threat(df, row_col="weed_science", label_col="weed_stage",
                                       color_col="type", hover_fn=hover,
                                       title="Weed Control Windows",
                                       stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                       row_label_map=row_label_map)
    detail_cols = ["weed_stage", "weed_science", "weed_name_en", "weed_name_th",
                   "common_name", "hrac_code", "type", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def insect_board_threat(crop_id, sheets, crop_stage_df, stage_label_col):
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

    fig = build_timeline_chart_threat(df, row_col="pest_name_en", label_col="pest_name_en",
                                       color_col="order", hover_fn=hover,
                                       title="Insect Pressure Windows",
                                       stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                       row_label_map=row_label_map)
    detail_cols = ["pest_name_en", "pest_name_th", "order", "common_name",
                   "irac_code", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def disease_board_threat(crop_id, sheets, crop_stage_df, stage_label_col):
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

    fig = build_timeline_chart_threat(df, row_col="disease_name_sc", label_col="disease_name_sc",
                                       color_col="type", hover_fn=hover,
                                       title="Disease Pressure Windows",
                                       stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                       row_label_map=row_label_map)
    detail_cols = ["disease_name_sc", "disease_name_en", "disease_name_th",
                   "common_name", "frac_code", "type", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def fertilizer_board_threat(crop_id, sheets, crop_stage_df, stage_label_col):
    is_thai = stage_label_col.endswith("_th")
    fert = sheets["fertilizer"]
    df = fert[fert["crop_id"] == crop_id].copy()

    has_stage = "stage" in df.columns
    has_type = "type" in df.columns and df["type"].notna().any()

    if has_type:
        type_options = sorted(df["type"].dropna().astype(str).unique().tolist())
        selected_types = st.multiselect(
            "Fertilizer type", type_options, default=type_options,
            help="Choose one type, or keep several selected to see them combined "
                 "on the same timeline (e.g. foliar + granular).",
            key="threat_fert_type",
        )
        df = df[df["type"].astype(str).isin(selected_types)] if selected_types else df.iloc[0:0]

    detail_cols = [c for c in ["stage", "type", "formula", "start_day", "end_day"]
                   if c in df.columns]
    detail_df = df[detail_cols].copy()

    if df.empty:
        fig = build_timeline_chart_threat(df, row_col="formula", label_col="formula",
                                           color_col="_none", hover_fn=lambda r: "",
                                           title="Fertilizer Application Windows",
                                           stage_df=crop_stage_df, stage_label_col=stage_label_col)
        return fig, detail_df, detail_cols

    row_col = "stage" if has_stage else "formula"
    if has_stage:
        stage_name_col = "stage_th" if (is_thai and "stage_th" in df.columns) else "stage"
        row_label_map = dict(zip(df[row_col], df[stage_name_col]))
    else:
        row_label_map = None

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

    fig = build_timeline_chart_threat(df_agg, row_col=row_col, label_col=row_col,
                                       color_col=color_col, hover_fn=hover,
                                       title="Fertilizer Application Windows",
                                       stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                       show_legend=has_type, row_label_map=row_label_map)
    return fig, detail_df, detail_cols


BOARDS_THREAT = {
    "Weed": weed_board_threat,
    "Insect": insect_board_threat,
    "Disease": disease_board_threat,
    "Fertilizer": fertilizer_board_threat,
}


def render_threat_view():
    st.title("🌾 Crop Threat & Input Dashboard")

    data_file = get_file_threat()
    if data_file is None:
        st.warning(
            f"No workbook found. Upload one from the sidebar, or place a file "
            f"named `{DEFAULT_PATH_THREAT}` next to `app.py`."
        )
        st.stop()

    try:
        sheets = load_workbook_threat(data_file)
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
        crop_choice = st.selectbox("Crop", list(crop_name_to_id.keys()), key="threat_crop")
    with col2:
        board_choice = st.selectbox("Board", list(BOARDS_THREAT.keys()), index=0, key="threat_board")
    with col3:
        stage_label_choice = st.radio("Label language", ["English", "Thai"],
                                       horizontal=True, key="threat_lang")
    label_col = "stage" if stage_label_choice == "English" else "stage_th"

    crop_id = crop_name_to_id[crop_choice]

    crop_stage_df = stage_df_all[stage_df_all["crop_id"] == crop_id]
    if crop_stage_df.empty:
        st.warning("No stage data for this crop.")
        st.stop()

    st.subheader(BOARD_TITLES_THREAT[board_choice])
    fig, board_df, detail_cols = BOARDS_THREAT[board_choice](crop_id, sheets, crop_stage_df, label_col)
    st.plotly_chart(fig, use_container_width=True)

    if board_df.empty:
        st.info(f"No {board_choice.lower()} data for this crop.")
    else:
        with st.expander(f"{board_choice} detail table (one row per product)"):
            st.dataframe(board_df, use_container_width=True, hide_index=True)


# =====================================================================
# =========================== COVERAGE VIEW ===========================
# Reads crop_timeline_coverage.xlsx:
#   crop_stage    : crop_id, crop, stage, stage_th, start_day, end_day
#   crop_weeds    : crop_id, ws_id, weed_stage, weed_id, weed_name_en,
#                   weed_name_th, weed_science, type, start_day, end_day
#   weed_her      : crop_id, ws_id, weed_id, weed_name_th, trade_name,
#                   company, common_name, concentration, formulation_type,
#                   hrac_code
#   crop_pest     : crop_id, pest_id, pest_name_en, pest_name_th, order,
#                   start_day, end_day
#   pest_ins      : crop_id, pest_id, pest_name_th, trade_name, company,
#                   common_name, concentration, formulation_type, irac_code
#   crop_disease  : crop_id, disease_id, disease_name_en, disease_name_th,
#                   disease_name_sc, type, start_day, end_day
#   disease_fun   : crop_id, disease_id, disease_name_th, trade_name,
#                   company, common_name, concentration, formulation_type,
#                   frac_code
#   crop_fer      : crop_id, crop, stage_id, stage, start_day, end_day
#   fertilizer    : crop_id, stage_id, formula, brand, company, stage, type
# =====================================================================

DEFAULT_PATH_COV = "crop_timeline_coverage.xlsx"

SHEET_NAMES_COV = [
    "crop_stage",
    "crop_weeds", "weed_her",
    "crop_pest", "pest_ins",
    "crop_disease", "disease_fun",
    "crop_fer", "fertilizer",
]

COVERED_COLOR = "#4CAF50"      # green — company has a product
NOT_COVERED_COLOR = "#E63946"  # red   — company has no product
COVERAGE_COLOR_MAP = {"Has Product": COVERED_COLOR, "No Product": NOT_COVERED_COLOR}

BOARD_TITLES_COV = {
    "Weed": "Weed Control Windows",
    "Insect": "Insect Pressure Windows",
    "Disease": "Disease Pressure Windows",
    "Fertilizer": "Fertilizer Application Windows",
}


@st.cache_data
def load_workbook_cov(file):
    sheets = {}
    for name in SHEET_NAMES_COV:
        try:
            df = pd.read_excel(file, sheet_name=name)
            df.columns = [c.strip() for c in df.columns]
            df = df.loc[:, ~df.columns.str.startswith("Unnamed:")]
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
            sheets[name] = df
        except ValueError:
            sheets[name] = pd.DataFrame()
    return sheets


def get_file_cov():
    st.sidebar.subheader("Product Coverage data source")
    uploaded = st.sidebar.file_uploader(
        "Upload workbook (.xlsx)", type=["xlsx"], key="cov_uploader"
    )
    if st.sidebar.button("🔄 Reload data", key="cov_reload"):
        st.cache_data.clear()
        st.rerun()
    if uploaded is not None:
        return uploaded
    if os.path.exists(DEFAULT_PATH_COV):
        return DEFAULT_PATH_COV
    return None


def _default_product_html(g: pd.DataFrame, code_col: str, code_label: str) -> str:
    lines = []
    for _, r in g.iterrows():
        trade = r.get("trade_name", "")
        common = r.get("common_name", "")
        conc = r.get("concentration", "")
        form = r.get("formulation_type", "")
        code = r.get(code_col, "")
        lines.append(f"• <b>{trade}</b> — {common} {conc} ({form}) [{code_label} {code}]")
    return "<br>".join(lines) if lines else "—"


def _fertilizer_product_html(g: pd.DataFrame, code_col: str, code_label: str) -> str:
    lines = []
    for _, r in g.iterrows():
        formula = r.get("formula", "")
        brand = r.get("brand", "")
        ftype = r.get("type", "")
        lines.append(f"• <b>{brand}</b> — {formula} ({ftype})")
    return "<br>".join(lines) if lines else "—"


def compute_coverage(window_df: pd.DataFrame, product_df: pd.DataFrame,
                      key_cols: list, company: str,
                      code_col: str, code_label: str,
                      product_html_fn=_default_product_html) -> pd.DataFrame:
    df = window_df.copy()

    if product_df.empty or not company:
        df["covered"] = False
        df["coverage_status"] = "No Product"
        df["product_list_html"] = "—"
        df["other_company_count"] = 0
        return df

    company_products = product_df[product_df["company"].astype(str) == str(company)]
    other_products = product_df[product_df["company"].astype(str) != str(company)]

    matched_map = {}
    if not company_products.empty:
        for keys, g in company_products.groupby(key_cols, dropna=False):
            matched_map[keys if isinstance(keys, tuple) else (keys,)] = product_html_fn(g, code_col, code_label)

    other_count_map = {}
    if not other_products.empty:
        for keys, g in other_products.groupby(key_cols, dropna=False):
            other_count_map[keys if isinstance(keys, tuple) else (keys,)] = g["company"].nunique()

    def _row_key(row):
        vals = tuple(row[c] for c in key_cols)
        return vals

    df["_key"] = df.apply(_row_key, axis=1)
    df["covered"] = df["_key"].isin(matched_map.keys())
    df["coverage_status"] = df["covered"].map({True: "Has Product", False: "No Product"})
    df["product_list_html"] = df["_key"].map(lambda k: matched_map.get(k, "—"))
    df["other_company_count"] = df["_key"].map(lambda k: other_count_map.get(k, 0))
    df = df.drop(columns=["_key"])
    return df


def get_companies_for_crop(sheets: dict, crop_id) -> list:
    companies = set()
    for name in ("weed_her", "pest_ins", "disease_fun", "fertilizer"):
        d = sheets.get(name, pd.DataFrame())
        if d.empty or "company" not in d.columns:
            continue
        d = d[d["crop_id"] == crop_id]
        companies.update(d["company"].dropna().astype(str).unique().tolist())
    return sorted(companies)


def build_timeline_chart_cov(df: pd.DataFrame, row_col: str,
                              color_col: str, hover_fn, title: str,
                              stage_df: pd.DataFrame = None, stage_label_col: str = "stage",
                              show_legend: bool = True, row_label_map: dict = None,
                              custom_color_map: dict = None,
                              force_show_legend: bool = False,
                              sort_col: str = None) -> go.Figure:
    if df.empty:
        fig = go.Figure()
        fig.update_layout(height=120, title=f"{title} — no data for this crop")
        return fig

    order_key = sort_col if sort_col else color_col
    order_df = (
        df.groupby(row_col)
        .agg(**{order_key: (order_key, "first"), "start_day": ("start_day", "min")})
        .reset_index()
        .sort_values([order_key, "start_day"])
    )
    row_order = order_df[row_col].tolist()
    row_to_base = {r: i for i, r in enumerate(row_order)}
    n_rows = len(row_order)

    color_values = sorted(df[color_col].dropna().astype(str).unique().tolist())
    color_map = dict(custom_color_map) if custom_color_map else {}
    default_palette = ["#457B9D", "#E76F51", "#2A9D8F", "#E9C46A", "#6A994E"]
    for i, v in enumerate(color_values):
        if v not in color_map:
            color_map[v] = default_palette[i % len(default_palette)]
    multi_category = len(color_values) > 1

    fig = go.Figure()
    annotations = []

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
            color = color_map.get(cat, "#999999")
            show_this_legend = (multi_category or force_show_legend) and cat not in seen_legend
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
    y_ticktext = [row_label_map.get(r, r) for r in row_order] if row_label_map else list(row_order)
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
        showlegend=(multi_category or force_show_legend) and show_legend,
        legend_title_text="Coverage",
        legend=dict(font=dict(size=17)),
        hoverlabel=dict(font=dict(size=20), align="left"),
        font=dict(size=17),
    )
    return fig


def weed_board_cov(crop_id, sheets, crop_stage_df, stage_label_col, company):
    is_thai = stage_label_col.endswith("_th")
    window_df = sheets["crop_weeds"][sheets["crop_weeds"]["crop_id"] == crop_id].copy()
    product_df = sheets["weed_her"][sheets["weed_her"]["crop_id"] == crop_id].copy()

    key_cols = ["ws_id", "weed_id"]
    df = compute_coverage(window_df, product_df, key_cols, company, "hrac_code", "HRAC")

    name_col = "weed_name_th" if is_thai else "weed_science"
    row_label_map = {
        r: f"[{t}] {n}" for r, n, t in zip(df["weed_science"], df[name_col], df["type"])
    }

    def hover(row):
        base = (
            f"<b><i>{row['weed_science']}</i></b><br>"
            f"{row['weed_name_en']} / {row['weed_name_th']}<br>"
            f"Type: {row.get('type', '')} | Stage: {row.get('weed_stage', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br><br>"
        )
        if row["covered"]:
            return base + f"<b>{company} products:</b><br>{row['product_list_html']}<extra></extra>"
        extra = f"<br><i>{row['other_company_count']} other company(ies) cover this</i>" if row['other_company_count'] else ""
        return base + f"<b>{company}: no product</b>{extra}<extra></extra>"

    fig = build_timeline_chart_cov(df, row_col="weed_science", color_col="coverage_status",
                                    hover_fn=hover, title="Weed Control Windows",
                                    stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                    row_label_map=row_label_map, sort_col="type",
                                    custom_color_map=COVERAGE_COLOR_MAP, force_show_legend=True)
    detail_cols = ["weed_stage", "weed_science", "weed_name_en", "weed_name_th",
                   "type", "start_day", "end_day", "coverage_status"]
    return fig, df[detail_cols], df["covered"].sum(), len(df)


def pest_board_cov(crop_id, sheets, crop_stage_df, stage_label_col, company):
    is_thai = stage_label_col.endswith("_th")
    window_df = sheets["crop_pest"][sheets["crop_pest"]["crop_id"] == crop_id].copy()
    product_df = sheets["pest_ins"][sheets["pest_ins"]["crop_id"] == crop_id].copy()

    key_cols = ["pest_id"]
    df = compute_coverage(window_df, product_df, key_cols, company, "irac_code", "IRAC")

    name_col = "pest_name_th" if is_thai else "pest_name_en"
    row_label_map = {
        r: f"[{o}] {n}" for r, n, o in zip(df["pest_name_en"], df[name_col], df["order"])
    }

    def hover(row):
        base = (
            f"<b>{row['pest_name_en']}</b><br>"
            f"{row['pest_name_th']}<br>"
            f"Insect order: {row.get('order', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br><br>"
        )
        if row["covered"]:
            return base + f"<b>{company} products:</b><br>{row['product_list_html']}<extra></extra>"
        extra = f"<br><i>{row['other_company_count']} other company(ies) cover this</i>" if row['other_company_count'] else ""
        return base + f"<b>{company}: no product</b>{extra}<extra></extra>"

    fig = build_timeline_chart_cov(df, row_col="pest_name_en", color_col="coverage_status",
                                    hover_fn=hover, title="Insect Pressure Windows",
                                    stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                    row_label_map=row_label_map, sort_col="order",
                                    custom_color_map=COVERAGE_COLOR_MAP, force_show_legend=True)
    detail_cols = ["pest_name_en", "pest_name_th", "order", "start_day", "end_day", "coverage_status"]
    return fig, df[detail_cols], df["covered"].sum(), len(df)


def disease_board_cov(crop_id, sheets, crop_stage_df, stage_label_col, company):
    is_thai = stage_label_col.endswith("_th")
    window_df = sheets["crop_disease"][sheets["crop_disease"]["crop_id"] == crop_id].copy()
    product_df = sheets["disease_fun"][sheets["disease_fun"]["crop_id"] == crop_id].copy()

    key_cols = ["disease_id"]
    df = compute_coverage(window_df, product_df, key_cols, company, "frac_code", "FRAC")

    name_col = "disease_name_th" if is_thai else "disease_name_en"
    row_label_map = {
        r: f"[{t}] {n}" for r, n, t in zip(df["disease_name_sc"], df[name_col], df["type"])
    }

    def hover(row):
        base = (
            f"<b><i>{row['disease_name_sc']}</i></b><br>"
            f"{row['disease_name_en']} / {row['disease_name_th']}<br>"
            f"Type: {row.get('type', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br><br>"
        )
        if row["covered"]:
            return base + f"<b>{company} products:</b><br>{row['product_list_html']}<extra></extra>"
        extra = f"<br><i>{row['other_company_count']} other company(ies) cover this</i>" if row['other_company_count'] else ""
        return base + f"<b>{company}: no product</b>{extra}<extra></extra>"

    fig = build_timeline_chart_cov(df, row_col="disease_name_sc", color_col="coverage_status",
                                    hover_fn=hover, title="Disease Pressure Windows",
                                    stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                    row_label_map=row_label_map, sort_col="type",
                                    custom_color_map=COVERAGE_COLOR_MAP, force_show_legend=True)
    detail_cols = ["disease_name_sc", "disease_name_en", "disease_name_th",
                   "type", "start_day", "end_day", "coverage_status"]
    return fig, df[detail_cols], df["covered"].sum(), len(df)


def fertilizer_board_cov(crop_id, sheets, crop_stage_df, stage_label_col, company):
    window_df = sheets["crop_fer"][sheets["crop_fer"]["crop_id"] == crop_id].copy()
    product_df_all = sheets["fertilizer"][sheets["fertilizer"]["crop_id"] == crop_id].copy()

    fert_types = sorted(product_df_all["type"].dropna().astype(str).unique().tolist()) \
        if "type" in product_df_all.columns else []
    type_choice = st.selectbox("Fertilizer type", ["All"] + fert_types, key="cov_fert_type")
    product_df = product_df_all if type_choice == "All" else \
        product_df_all[product_df_all["type"].astype(str) == type_choice]

    key_cols = ["stage_id"]
    df = compute_coverage(window_df, product_df, key_cols, company, "type", "Type",
                           product_html_fn=_fertilizer_product_html)

    row_label_map = dict(zip(df["stage_id"], df["stage"]))

    def hover(row):
        base = f"<b>{row['stage']}</b><br>Day {row['start_day']}–{row['end_day']}<br><br>"
        if row["covered"]:
            return base + f"<b>{company} products:</b><br>{row['product_list_html']}<extra></extra>"
        extra = f"<br><i>{row['other_company_count']} other company(ies) cover this</i>" if row['other_company_count'] else ""
        return base + f"<b>{company}: no product</b>{extra}<extra></extra>"

    fig = build_timeline_chart_cov(df, row_col="stage_id", color_col="coverage_status",
                                    hover_fn=hover, title="Fertilizer Application Windows",
                                    stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                    row_label_map=row_label_map, sort_col="start_day",
                                    custom_color_map=COVERAGE_COLOR_MAP, force_show_legend=True)
    detail_cols = ["stage", "start_day", "end_day", "coverage_status"]
    return fig, df[detail_cols], df["covered"].sum(), len(df)


BOARDS_COV = {
    "Weed": weed_board_cov,
    "Insect": pest_board_cov,
    "Disease": disease_board_cov,
    "Fertilizer": fertilizer_board_cov,
}


def render_coverage_view():
    st.title("🧭 Product Coverage Dashboard")
    st.caption("Pick a company to see where it has products (green) and where it's lagging (red).")

    data_file = get_file_cov()
    if data_file is None:
        st.warning(
            f"No workbook found. Upload one from the sidebar, or place a file "
            f"named `{DEFAULT_PATH_COV}` next to `app.py`."
        )
        st.stop()

    try:
        sheets = load_workbook_cov(data_file)
    except Exception as e:
        st.error(f"Couldn't read the workbook: {e}")
        st.stop()

    stage_df_all = sheets["crop_stage"]
    if stage_df_all.empty:
        st.error("`crop_stage` sheet is missing or empty.")
        st.stop()

    crop_lookup = stage_df_all[["crop_id", "crop"]].drop_duplicates()
    crop_name_to_id = dict(zip(crop_lookup["crop"], crop_lookup["crop_id"]))

    col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
    with col1:
        crop_choice = st.selectbox("Crop", list(crop_name_to_id.keys()), key="cov_crop")
    crop_id = crop_name_to_id[crop_choice]

    companies = get_companies_for_crop(sheets, crop_id)
    with col2:
        if companies:
            company_choice = st.selectbox("Company", companies, key="cov_company")
        else:
            company_choice = None
            st.warning("No companies found for this crop across weed_her / pest_ins / disease_fun / fertilizer.")
    with col3:
        board_choice = st.selectbox("Board", list(BOARDS_COV.keys()), index=0, key="cov_board")
    with col4:
        stage_label_choice = st.radio("Label language", ["English", "Thai"],
                                       horizontal=True, key="cov_lang")
    label_col = "stage" if stage_label_choice == "English" else "stage_th"

    crop_stage_df = stage_df_all[stage_df_all["crop_id"] == crop_id]
    if crop_stage_df.empty:
        st.warning("No stage data for this crop.")
        st.stop()

    st.subheader(f"{BOARD_TITLES_COV[board_choice]} — {company_choice or 'no company selected'}")

    if company_choice:
        fig, detail_df, _covered_n, _total_n = BOARDS_COV[board_choice](
            crop_id, sheets, crop_stage_df, label_col, company_choice
        )
        st.plotly_chart(fig, use_container_width=True)

        if detail_df.empty:
            st.info(f"No {board_choice.lower()} data for this crop.")
        else:
            with st.expander(f"{board_choice} detail table"):
                st.dataframe(detail_df, use_container_width=True, hide_index=True)
    else:
        st.info("Select a company above to see its coverage.")


# =====================================================================
# App — sidebar view switch
# =====================================================================

st.sidebar.subheader("View")
view = st.sidebar.radio(
    "Choose a dashboard",
    ["Crop Threat & Input", "Product Coverage"],
    label_visibility="collapsed",
    key="view_switch",
)

if view == "Crop Threat & Input":
    render_threat_view()
else:
    render_coverage_view()
