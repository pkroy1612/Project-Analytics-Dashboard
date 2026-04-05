import os
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit_autorefresh import st_autorefresh

# ==============================
# Page config
# ==============================
st.set_page_config(
    page_title="Project Analytics Dashboard",
    layout="wide",
    page_icon="📊"
)

# ==============================
# Sidebar
# ==============================
st.sidebar.header("📂 Data Source")
path = st.sidebar.text_input("Excel file path (.xlsx)", value="sample.xlsx")
sheet_name = st.sidebar.text_input("Sheet name (optional)", value="")

st.sidebar.header("⏳ Lag Definition")
today = st.sidebar.date_input("Today", value=date.today())

st.sidebar.header("🎨 Visualization")
color_by = st.sidebar.selectbox("Color Gantt by", ["vendor", "status"])

# ==============================
# Helper functions
# ==============================
def normalize_columns(df):
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df

def coerce_percent(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, str):
        x = x.replace("%", "").strip()
    try:
        return float(x)
    except:
        return np.nan

def load_excel(file_path, sheet):
    if not os.path.exists(file_path):
        raise FileNotFoundError("Excel file not found")

    df = pd.read_excel(file_path, sheet_name=(sheet if sheet else 0))
    df = normalize_columns(df)

    if "%complete" not in df.columns:
        raise ValueError("Required column '%complete' not found")

    df.rename(columns={"%complete": "pct_complete"}, inplace=True)

    required = ["vendor", "project", "activity", "start", "end", "pct_complete"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Missing column: {col}")

    df["vendor"] = df["vendor"].ffill()
    df["project"] = df["project"].ffill()

    df["activity"] = df["activity"].astype(str)
    df["start"] = pd.to_datetime(df["start"])
    df["end"] = pd.to_datetime(df["end"])
    df["pct_complete"] = df["pct_complete"].apply(coerce_percent).fillna(0).clip(0, 100)

    return df

def add_metrics(df, today_dt):
    df = df.copy()
    df["duration_days"] = (df["end"] - df["start"]).dt.days.clip(lower=0)
    df["actual_progress"] = df["pct_complete"] / 100

    planned = (today_dt - df["start"]) / (df["end"] - df["start"])
    df["planned_progress"] = planned.clip(0, 1)

    df["behind_by_pct"] = (df["planned_progress"] - df["actual_progress"]) * 100

    df["overdue_days"] = np.where(
        (df["pct_complete"] < 100) & (today_dt > df["end"]),
        (today_dt - df["end"]).dt.days,
        0
    )

    df["status"] = np.select(
        [
            df["pct_complete"] >= 100,
            df["overdue_days"] > 0,
            df["behind_by_pct"] > 5
        ],
        ["Completed", "Overdue", "Behind"],
        default="On track"
    )
    return df

# ==============================
# Main
# ==============================
st.title("📊 Project Analytics Dashboard")

try:
    df_raw = load_excel(path, sheet_name.strip() or None)
    df = add_metrics(df_raw, pd.Timestamp(today))

    # ==============================
    # 🔽 PROJECT DROPDOWN
    # ==============================
    projects = ["All Projects"] + sorted(df["project"].unique().tolist())
    selected_project = st.sidebar.selectbox("📁 Select Project", projects)

    if selected_project != "All Projects":
        df = df[df["project"] == selected_project]

    if df.empty:
        st.warning("No data for selected project")
        st.stop()

    # ==============================
    # KPIs (PROJECT-SPECIFIC)
    # ==============================
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Activities", len(df))
    k2.metric("Overdue Activities", int((df["overdue_days"] > 0).sum()))
    k3.metric("Behind Schedule", int((df["status"] == "Behind").sum()))
    k4.metric("Max Overdue (days)", int(df["overdue_days"].max()))

    st.divider()

    tab1, tab2, tab3 = st.tabs(
        ["📊 Gantt Chart", "⚠️ Delays", "📌 Project Details"]
    )

    # ==============================
    # GANTT
    # ==============================
    with tab1:
        fig = px.timeline(
            df,
            x_start="start",
            x_end="end",
            y="activity",
            color=color_by,
            hover_data=["vendor", "pct_complete", "status"],
            template="plotly_white"
        )
        fig.update_yaxes(autorange="reversed")
        st.plotly_chart(fig, use_container_width=True)

    # ==============================
    # RISKS
    # ==============================
    with tab2:
        overdue = df[df["overdue_days"] > 0]
        if overdue.empty:
            st.info("No overdue activities")
        else:
            fig = px.bar(
                overdue,
                x="overdue_days",
                y="activity",
                orientation="h",
                template="plotly_white"
            )
            st.plotly_chart(fig, use_container_width=True)

    # ==============================
    # PROJECT DETAILS TABLE
    # ==============================
    with tab3:
        st.dataframe(
            df[
                [
                    "vendor", "activity", "start", "end",
                    "pct_complete", "status", "overdue_days"
                ]
            ].sort_values("start"),
            use_container_width=True
        )

except Exception as e:
    st.error(str(e))
