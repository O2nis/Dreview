"""S‑Curve & Document Tracking Dashboard
--------------------------------------------------
Streamlit one‑file app that:
• plots an S‑curve (expected vs actual cumulative documents)
• **NEW:** bar chart of cumulative *Actual* vs *Expected* grouped by *Area*
• **NEW:** donut chart showing *Final (Flag = 1)* vs *Total* documents
• stacked‑bar of documents by *Milestone* **ignoring Flag** (all docs counted)

Assumptions on column names (edit if yours differ)
--------------------------------------------------
PlannedDate … Planned submission date (datetime)
ActualDate  … Actual submission date (datetime)
Area        … Discipline / area code (str)
Milestone   … Reply‑by‑EPC / contractual step (str)
Flag        … 1 if document is “Final”, 0 otherwise (int)

Dependencies: streamlit, pandas, matplotlib, numpy
Run:  
```bash
streamlit run s_curve_app.py
```
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime as dt

plt.rcParams.update({"font.size": 9})

###############################################################################
# Helpers
###############################################################################

def date_cols_to_datetime(df, cols):
    """Parse the listed columns to datetime (in‑place)."""
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

###############################################################################
# Main App
###############################################################################

def main():
    st.title("📈 S‑Curve & Document Tracking Dashboard")

    uploaded = st.file_uploader("Upload the tracking CSV / XLSX", type=["csv", "xls", "xlsx"])
    if uploaded is None:
        st.info("⬆️ Upload a file to get started.")
        st.stop()

    # ----------------------------------------------------------------------
    # LOAD DATA
    # ----------------------------------------------------------------------
    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)

    # Basic tidy‑ups ---------------------------------------------------------
    df = date_cols_to_datetime(df, ["PlannedDate", "ActualDate"])

    # Provide simple stats at the top --------------------------------------
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Total documents", len(df))
    if "Flag" in df.columns:
        col_b.metric("Final (Flag = 1)", int(df["Flag"].eq(1).sum()))
    if "ActualDate" in df.columns:
        col_c.metric("Docs delivered", int(df["ActualDate"].notna().sum()))

    # ----------------------------------------------------------------------
    # S‑Curve (cumulative Expected vs Actual) ------------------------------
    # ----------------------------------------------------------------------
    if {"PlannedDate", "ActualDate"}.issubset(df.columns):
        plan_curve = (
            df.groupby("PlannedDate").size().cumsum().rename("Expected")
        )
        act_curve = (
            df.groupby("ActualDate").size().cumsum().rename("Actual")
        )
        curve = (
            pd.concat([plan_curve, act_curve], axis=1)
            .fillna(method="ffill")
            .fillna(0)
        )

        fig, ax = plt.subplots()
        ax.plot(curve.index, curve["Expected"], label="Expected", lw=2)
        ax.plot(curve.index, curve["Actual"], label="Actual", lw=2)
        ax.set_ylabel("Cumulative documents")
        ax.set_xlabel("")
        ax.legend()
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(ax.xaxis.get_major_locator()))
        st.pyplot(fig)
    else:
        st.warning("Columns 'PlannedDate' and 'ActualDate' are required for the S‑curve.")

    # ----------------------------------------------------------------------
    # NEW 1️⃣  – Bar: Actual vs Expected grouped by Area --------------------
    # ----------------------------------------------------------------------
    if {"Area", "PlannedDate", "ActualDate"}.issubset(df.columns):
        area_stats = (
            df.groupby("Area").agg(Expected=("PlannedDate", "size"), Actual=("ActualDate", "count"))
        ).reset_index()

        fig2, ax2 = plt.subplots()
        x = np.arange(len(area_stats))
        width = 0.35
        ax2.bar(x - width / 2, area_stats["Expected"], width, label="Expected")
        ax2.bar(x + width / 2, area_stats["Actual"], width, label="Actual")
        ax2.set_xticks(x)
        ax2.set_xticklabels(area_stats["Area"], rotation=45, ha="right")
        ax2.set_ylabel("Cumulative documents")
        ax2.set_title("Cumulative Actual vs Expected by Area")
        ax2.legend()
        st.pyplot(fig2)
    else:
        st.info("Need 'Area', 'PlannedDate', 'ActualDate' columns for the area bar chart.")

    # ----------------------------------------------------------------------
    # NEW 2️⃣  – Donut: Final vs Total --------------------------------------
    # ----------------------------------------------------------------------
    if "Flag" in df.columns:
        final_docs = int(df["Flag"].eq(1).sum())
        total_docs = len(df)
        other_docs = total_docs - final_docs

        fig3, ax3 = plt.subplots()
        wedges, *_ = ax3.pie(
            [final_docs, other_docs],
            labels=["Final (Flag=1)", "Other"],
            autopct="%1.0f%%",
            startangle=90,
            wedgeprops=dict(width=0.4),
        )
        ax3.set_aspect("equal")
        ax3.set_title("Final vs Total Documents")
        st.pyplot(fig3)
    else:
        st.info("Column 'Flag' not found – skipping Final vs Total donut.")

    # ----------------------------------------------------------------------
    # Stacked bar by Milestone (ignoring Flag) ------------------------------
    # ----------------------------------------------------------------------
    if "Milestone" in df.columns:
        milestone_counts = df["Milestone"].value_counts().sort_index()
        fig4, ax4 = plt.subplots()
        ax4.bar(milestone_counts.index, milestone_counts.values)
        ax4.set_ylabel("Documents")
        ax4.set_title("Documents by Milestone (all flags)")
        ax4.tick_params(axis="x", rotation=45)
        st.pyplot(fig4)
    else:
        st.info("Column 'Milestone' not present – milestone chart skipped.")


if __name__ == "__main__":
    main()
