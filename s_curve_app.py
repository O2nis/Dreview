import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime as dt
import numpy as np
import seaborn as sns
from cycler import cycler
import uuid

plt.rcParams.update({'font.size': 8})

def parse_date(d):
    if not isinstance(d, str):
        return pd.NaT
    d = d.strip()
    if d in ['0-Jan-00', '########', '']:
        return pd.NaT
    try:
        return pd.to_datetime(d, format='%d-%b-%y', errors='coerce')
    except:
        return pd.NaT

def main():
    st.set_page_config(page_title="S-Curve Analysis", layout="wide")

    # ----------------------------------------------------------------
    # A) TEMPLATE CSV DOWNLOAD (with updated columns)
    # ----------------------------------------------------------------
    TEMPLATE_CONTENT = """ID,Discipline,Area,Document Title,Project Identifier,Originator,Document Number,Document Type,Counter,Revision,Area Code,Disc,Category,Transmittal Code,Comment Sheet OE,Comment Sheet EPC,Schedule [Days],Issued By EPC,Issuance Expected,Review By OE,Expected Review,Reply By EPC,Final Issuance Expected,Man Hours,Status,CS rev,Flag
1,Piping,Area1,AAA,Proj1,Orig1,Doc00001,Type1,Cnt1,Rev0,AC1,D1,Cat1,TC1,CSOE1,CSEPC1,45,8-Oct-24,,8-Dec-24,,,80,APP,Rev1,1
2,Electrical,Area2,BBB,Proj2,Orig2,Doc00002,Type2,Cnt2,Rev1,AC2,D2,Cat2,TC2,CSOE2,CSEPC2,30,8-Mar-25,,,,,100,REJ,Rev2,0
3,Instrument,Area3,CCC,Proj3,Orig3,Doc00003,Type3,Cnt3,Rev2,AC3,D3,Cat3,TC3,CSOE3,CSEPC3,60,8-Mar-25,,,,,120,Closed,Rev3,1
4,Economic,Area4,CCC,Proj4,Orig4,Doc00003,Type4,Cnt4,Rev2,AC4,D4,Cat4,TC4,CSOE4,CSEPC4,360,,,,,,220,,Rev4,0
"""

    st.subheader("Download CSV Template")
    st.download_button(
        label="Download Template CSV",
        data=TEMPLATE_CONTENT.encode("utf-8"),
        file_name="EDDR_template_updated.csv",
        mime="text/csv"
    )

    # --------------------------
    # 1) SIDEBAR INPUTS
    # --------------------------
    st.sidebar.header("Configuration")
    CSV_INPUT_PATH = st.sidebar.file_uploader("Upload Input CSV", type=["csv"])
    INITIAL_DATE = st.sidebar.date_input("Initial Date for Expected Calculations", value=pd.to_datetime("2024-08-01"))
    ISSUED_WEIGHT = st.sidebar.number_input("Issued By EPC Weight", value=0.40, step=0.05)
    REVIEW_WEIGHT = st.sidebar.number_input("Review By OE Weight", value=0.30, step=0.05)
    REPLY_WEIGHT = st.sidebar.number_input("Reply By EPC Weight", value=0.30, step=0.05)
    RECOVERY_FACTOR = st.sidebar.number_input("Recovery Factor", value=0.75, step=0.05)

    REVIEW_DELTA_DAYS = st.sidebar.number_input("Days to add for Expected Review", value=10, step=1)
    REPLY_DELTA_DAYS = st.sidebar.number_input("Days to add for Final Issuance Expected", value=5, step=1)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Visualization Settings")
    seaborn_style = st.sidebar.selectbox(
        "Select Seaborn Style",
        ["darkgrid", "whitegrid", "dark", "white", "ticks"],
        index=1
    )
    sns.set_style(seaborn_style)

    display_unit = st.sidebar.selectbox(
        "Y-Axis Unit for Man-Hours Charts",
        ["Man-Hours", "Percentage"],
        index=0
    )

    st.sidebar.markdown("### S-Curve Color Scheme")
    actual_color = st.sidebar.color_picker("Actual Progress Color", "#1f77b4")
    expected_color = st.sidebar.color_picker("Expected Progress Color", "#ff7f0e")
    projected_color = st.sidebar.color_picker("Projected Recovery Color", "#2ca02c")
    today_color = st.sidebar.color_picker("Today Line Color", "#000000")
    end_date_color = st.sidebar.color_picker("End Date Line Color", "#d62728")

    st.sidebar.markdown("### Bar/Donut Color Scheme")
    color_choice = st.sidebar.selectbox(
        "Select Color Scheme (applies to bar/donut charts):",
        ["Standard", "Shades of Blue", "Shades of Green"]
    )

    if CSV_INPUT_PATH is None:
        st.warning("Please upload your input CSV file or download the template above.")
        return

    # --------------------------
    # 2) LOAD CSV & PREP DATA
    # --------------------------
    df = pd.read_csv(CSV_INPUT_PATH)
    df.columns = [
        "ID", "Discipline", "Area", "Document Title", "Project Identifier", "Originator",
        "Document Number", "Document Type", "Counter", "Revision", "Area Code", "Disc",
        "Category", "Transmittal Code", "Comment Sheet OE", "Comment Sheet EPC",
        "Schedule [Days]", "Issued By EPC", "Issuance Expected", "Review By OE",
        "Expected Review", "Reply By EPC", "Final Issuance Expected", "Man Hours",
        "Status", "CS rev", "Flag"
    ]

    df["Issued By EPC"] = df["Issued By EPC"].apply(parse_date)
    df["Review By OE"] = df["Review By OE"].apply(parse_date)
    df["Reply By EPC"] = df["Reply By EPC"].apply(parse_date)

    df["Schedule [Days]"] = pd.to_numeric(df["Schedule [Days]"], errors="coerce").fillna(0)
    df["Man Hours"] = pd.to_numeric(df["Man Hours"], errors="coerce").fillna(0)
    df["Flag"] = pd.to_numeric(df["Flag"], errors="coerce").fillna(0).astype(int)

    df["Issuance Expected"] = pd.Timestamp(INITIAL_DATE) + pd.to_timedelta(df["Schedule [Days]"], unit="D")
    df["Expected Review"] = df["Issuance Expected"] + dt.timedelta(days=REVIEW_DELTA_DAYS)
    df["Final Issuance Expected"] = df["Expected Review"] + dt.timedelta(days=REPLY_DELTA_DAYS)
    df["Final Issuance Expected"] = pd.to_datetime(df["Final Issuance Expected"], errors='coerce')

    final_issuance_max = df["Final Issuance Expected"].dropna().max()
    if pd.isna(final_issuance_max):
        st.warning("No valid Final Issuance Expected in data. Falling back to overall max date.")
        final_issuance_max = (
            pd.Series(df[["Issuance Expected", "Expected Review", "Final Issuance Expected"]].values.ravel())
            .dropna()
            .max()
        )

    date_cols = ["Issued By EPC", "Review By OE", "Reply By EPC", "Issuance Expected", "Expected Review", "Final Issuance Expected"]
    all_dates = df[date_cols].values.ravel()
    valid_dates = pd.Series(all_dates).dropna()
    if valid_dates.empty:
        st.error("No valid milestone dates found.")
        return

    start_date = valid_dates.min()
    today_date = pd.to_datetime("today").normalize()

    total_man_hours = df["Man Hours"].sum()

    # --------------------------
    # 3) BUILD TIMELINES
    # --------------------------
    actual_timeline = pd.date_range(start=start_date, end=today_date, freq='W')
    expected_timeline = pd.date_range(start=start_date, end=final_issuance_max, freq='W')

    # --------------------------
    # 4) BUILD ACTUAL AND EXPECTED CUMULATIVE VALUES
    # --------------------------
    actual_cum = []
    last_actual_value = 0.0
    last_progress_date = start_date

    for current_date in actual_timeline:
        a_sum = 0.0
        has_progress = False
        for _, row in df.iterrows():
            mh = row["Man Hours"]
            a_val = 0.0
            if pd.notna(row["Issued By EPC"]) and row["Issued By EPC"] <= current_date:
                a_val += ISSUED_WEIGHT
            if pd.notna(row["Review By OE"]) and row["Review By OE"] <= current_date:
                a_val += REVIEW_WEIGHT
            if pd.notna(row["Reply By EPC"]) and row["Reply By EPC"] <= current_date and row["Flag"] == 1:
                a_val += REPLY_WEIGHT
            if a_val > 0:
                has_progress = True
            a_sum += mh * a_val

        if has_progress:
            last_actual_value = a_sum
            last_progress_date = current_date
            actual_cum.append(a_sum)
        else:
            actual_cum.append(last_actual_value)

    if last_progress_date < today_date:
        actual_timeline = list(actual_timeline) + [today_date]
        actual_cum = actual_cum + [last_actual_value]

    expected_cum = []
    last_expected_value = 0.0
    last_expected_progress_date = start_date

    for current_date in expected_timeline:
        e_sum = 0.0
        has_progress = False
        for _, row in df.iterrows():
            mh = row["Man Hours"]
            e_val = 0.0
            if pd.notna(row["Issuance Expected"]) and row["Issuance Expected"] <= current_date:
                e_val += ISSUED_WEIGHT
            if pd.notna(row["Expected Review"]) and row["Expected Review"] <= current_date:
                e_val += REVIEW_WEIGHT
            if pd.notna(row["Final Issuance Expected"]) and row["Final Issuance Expected"] <= current_date:
                e_val += REPLY_WEIGHT
            if e_val > 0:
                has_progress = True
            e_sum += mh * e_val

        if has_progress:
            last_expected_value = e_sum
            last_expected_progress_date = current_date
            expected_cum.append(e_sum)
        else:
            expected_cum.append(last_expected_value)

    if pd.notna(final_issuance_max) and last_expected_progress_date < final_issuance_max:
        expected_timeline = list(expected_timeline) + [final_issuance_max]
        expected_cum = expected_cum + [last_expected_value]

    final_actual = actual_cum[-1]
    final_expected = expected_cum[-1]

    # Convert to percentage if selected
    actual_cum_display = [(x / total_man_hours * 100) if display_unit == "Percentage" and total_man_hours > 0 else x for x in actual_cum]
    expected_cum_display = [(x / total_man_hours * 100) if display_unit == "Percentage" and total_man_hours > 0 else x for x in expected_cum]
    final_actual_display = (final_actual / total_man_hours * 100) if display_unit == "Percentage" and total_man_hours > 0 else final_actual
    final_expected_display = (final_expected / total_man_hours * 100) if display_unit == "Percentage" and total_man_hours > 0 else final_expected

    # --------------------------
    # 5) PROJECTED RECOVERY LINE
    # --------------------------
    if today_date <= actual_timeline[0]:
        today_idx = 0
    elif today_date >= actual_timeline[-1]:
        today_idx = len(actual_timeline) - 1
    else:
        today_idx = np.searchsorted(actual_timeline, today_date, side="right") - 1

    actual_today = actual_cum[today_idx]
    actual_today_display = (actual_today / total_man_hours * 100) if display_unit == "Percentage" and total_man_hours > 0 else actual_today

    if today_date <= expected_timeline[0]:
        expected_today_idx = 0
    elif today_date >= expected_timeline[-1]:
        expected_today_idx = len(expected_timeline) - 1
    else:
        expected_today_idx = np.searchsorted(expected_timeline, today_date, side="right") - 1
    expected_today = expected_cum[expected_today_idx]
    expected_today_display = (expected_today / total_man_hours * 100) if display_unit == "Percentage" and total_man_hours > 0 else expected_today

    gap_hrs = final_expected - actual_today

    projected_timeline = []
    projected_cumulative = []
    recovery_end_date = None

    if gap_hrs > 0:
        total_days_span = (final_issuance_max - start_date).days
mu
        project_months = total_days_span / 30.4
        delay_fraction = gap_hrs / final_expected if final_expected > 0 else 0
        T_recover_months = project_months * delay_fraction * RECOVERY_FACTOR
        T_recover_weeks = T_recover_months * (30.4 / 7.0)

        if T_recover_weeks < 1:
            T_recover_weeks = 1

        slope_new = gap_hrs / T_recover_weeks
        last_date = today_date
        cum_val = actual_today

        steps = int(T_recover_weeks) + 2
        for _ in range(steps):
            projected_timeline.append(last_date)
            cum_val_display = (cum_val / total_man_hours * 100) if display_unit == "Percentage" and total_man_hours > 0 else cum_val
            projected_cumulative.append(cum_val_display)
            if cum_val >= final_expected:
                break
            last_date = last_date + dt.timedelta(weeks=1)
            cum_val = min(final_expected, cum_val + slope_new)

        recovery_end_date = last_date

    # --------------------------
    # 6) S-CURVE
    # --------------------------
    st.subheader("S-Curve with Delay Recovery")
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(actual_timeline, actual_cum_display, label="Actual Progress", color=actual_color, linewidth=2)
    ax.plot(expected_timeline, expected_cum_display, label="Expected Progress", color=expected_color, linewidth=2)

    if last_progress_date < today_date:
        ax.hlines(y=actual_cum_display[-1], xmin=last_progress_date, xmax=today_date,
                 color=actual_color, linestyle='-', linewidth=2)

    if pd.notna(final_issuance_max) and last_expected_progress_date < final_issuance_max:
        ax.hlines(y=expected_cum_display[-1], xmin=last_expected_progress_date, xmax=final_issuance_max,
                 color=expected_color, linestyle='-', linewidth=2)

    if projected_timeline:
        ax.plot(
            projected_timeline,
            projected_cumulative,
            linestyle=":",
            label="Projected (Recovery Factor)",
            color=projected_color,
            linewidth=3
        )

    ax.set_title("S-Curve with Delay Recovery", fontsize=12)
    ax.set_xlabel("Date", fontsize=10)
    ax.set_ylabel("Cumulative Progress (%s)" % ("%" if display_unit == "Percentage" else "Man-Hours"), fontsize=10)
    ax.grid(True)
    ax.legend(fontsize=9)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b-%Y"))
    plt.xticks(rotation=45)
    plt.tight_layout()

    ax.axvline(today_date, color=today_color, linestyle="--", linewidth=1.5, label="Today")
    ax.annotate(
        f"Today\n{today_date.strftime('%d-%b-%Y')}",
        xy=(today_date, final_expected_display * 0.1),
        xytext=(10, 10),
        textcoords="offset points",
        color=today_color,
        bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.7),
        fontsize=8
    )

    if pd.notna(final_issuance_max):
        ax.axvline(final_issuance_max, linestyle="--", linewidth=1.5, color=end_date_color)
        ax.annotate(
            f"Original End\n{final_issuance_max.strftime('%d-%b-%Y')}",
            xy=(final_issuance_max, final_expected_display * 0.2),
            xytext=(-100, 10),
            textcoords="offset points",
            color=end_date_color,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.7),
            fontsize=8,
            arrowprops=dict(arrowstyle="->", color=end_date_color)
        )

    if recovery_end_date:
        ax.axvline(recovery_end_date, linestyle="--", linewidth=1.5, color=end_date_color)
        ax.annotate(
            f"Recovery End\n{recovery_end_date.strftime('%d-%b-%Y')}",
            xy=(recovery_end_date, final_expected_display * 0.3),
            xytext=(10, 10),
            textcoords="offset points",
            color=end_date_color,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.7),
            fontsize=8,
            arrowprops=dict(arrowstyle="->", color=end_date_color)
        )

    delay_today = expected_today - actual_today
    delay_display = (delay_today / total_man_hours * 100) if display_unit == "Percentage" and total_man_hours > 0 else delay_today
    ax.annotate(
        f"Current Delay: {delay_display:,.1f} {'%' if display_unit == 'Percentage' else 'MH'}\n({delay_today/final_expected*100:.1f}%)",
        xy=(today_date, (actual_today_display + expected_today_display)/2),
        xytext=(10, -50),
        textcoords="offset points",
        color=today_color,
        bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.7),
        arrowprops=dict(arrowstyle="->", color=today_color),
        ha="left",
        fontsize=8
    )

    st.pyplot(fig)

    # --------------------------
    # 7) COLOR SCHEME FOR OTHER CHARTS
    # --------------------------
    standard_cycler = plt.rcParamsDefault['axes.prop_cycle']
    blue_cycler = cycler(color=["#cce5ff", "#99ccff", "#66b2ff", "#3399ff", "#007fff"])
    green_cycler = cycler(color=["#ccffcc", "#99ff99", "#66ff66", "#33cc33", "#009900"])

    if color_choice == "Shades of Blue":
        plt.rc("axes", prop_cycle=blue_cycler)
    elif color_choice == "Shades of Green":
        plt.rc("axes", prop_cycle=green_cycler)
    else:
        plt.rc("axes", prop_cycle=standard_cycler)

    # --------------------------
    # 8) ACTUAL vs EXPECTED PROGRESS BY DISCIPLINE
    # --------------------------
    end_date = max(today_date, final_issuance_max)
    final_actual_progress = []
    final_expected_progress = []

    for _, row in df.iterrows():
        a_prog = 0.0
        if pd.notna(row["Issued By EPC"]) and row["Issued By EPC"] <= end_date:
            a_prog += ISSUED_WEIGHT
        if pd.notna(row["Review By OE"]) and row["Review By OE"] <= end_date:
            a_prog += REVIEW_WEIGHT
        if pd.notna(row["Reply By EPC"]) and row["Reply By EPC"] <= end_date and row["Flag"] == 1:
            a_prog += REPLY_WEIGHT

        e_prog = 0.0
        if pd.notna(row["Issuance Expected"]) and row["Issuance Expected"] <= end_date:
            e_prog += ISSUED_WEIGHT
        if pd.notna(row["Expected Review"]) and row["Expected Review"] <= end_date:
            e_prog += REVIEW_WEIGHT
        if pd.notna(row["Final Issuance Expected"]) and row["Final Issuance Expected"] <= end_date:
            e_prog += REPLY_WEIGHT

        final_actual_progress.append(row["Man Hours"] * a_prog)
        final_expected_progress.append(row["Man Hours"] * e_prog)

    df["Actual_Progress_At_Final"] = final_actual_progress
    df["Expected_Progress_At_Final"] = final_expected_progress

    by_disc = df.groupby("Discipline")[["Actual_Progress_At_Final", "Expected_Progress_At_Final"]].sum()

    if display_unit == "Percentage" and total_man_hours > 0:
        by_disc["Actual_Progress_At_Final"] = by_disc["Actual_Progress_At_Final"] / total_man_hours * 100
        by_disc["Expected_Progress_At_Final"] = by_disc["Expected_Progress_At_Final"] / total_man_hours * 100

    st.subheader("Actual vs. Expected Progress by Discipline")
    fig2, ax2 = plt.subplots(figsize=(8, 5))
    x = range(len(by_disc.index))
    width = 0.35

    ax2.bar(
        [i - width/2 for i in x],
        by_disc["Actual_Progress_At_Final"],
        width=width,
        label='Actual Progress'
    )
    ax2.bar(
        [i + width/2 for i in x],
        by_disc["Expected_Progress_At_Final"],
        width=width,
        label='Expected Progress'
    )

    ax2.set_title("Actual vs. Expected Progress by Discipline", fontsize=10)
    ax2.set_xlabel("Discipline", fontsize=9)
    ax2.set_ylabel("Cumulative Progress (%s)" % ("%" if display_unit == "Percentage" else "Man-Hours"), fontsize=9)
    ax2.set_xticks(ticks=x)
    ax2.set_xticklabels(by_disc.index, rotation=45, ha='right', fontsize=8)
    ax2.legend(fontsize=8)
    plt.tight_layout()
    st.pyplot(fig2)

    # --------------------------
    # 9) TWO DONUT CHARTS SIDE BY SIDE
    # --------------------------
    total_actual = df["Actual_Progress_At_Final"].sum()
    overall_pct = total_actual / total_man_hours if total_man_hours > 0 else 0

    issued_delivered = df["Issued By EPC"].notna().sum()
    total_docs = len(df)
    issued_values = [issued_delivered, total_docs - issued_delivered]

    st.subheader("Donut Charts")
    col1, col2 = st.columns(2)

    with col1:
        st.write("**Total Project Completion**")
        fig3, ax3 = plt.subplots(figsize=(4, 4))
        ax3.pie(
            [overall_pct, 1 - overall_pct],
            labels=["Completed", "Remaining"],
            autopct="%1.1f%%",
            startangle=140,
            wedgeprops={"width": 0.4}
        )
        ax3.set_title("Project Completion", fontsize=9)
        st.pyplot(fig3)

    with col2:
        st.write("**Issued By EPC Delivery Status**")
        def issued_autopct(pct):
            total_count = sum(issued_values)
            docs = int(round(pct * total_count / 100.0))
            return f"{docs} docs" if docs > 0 else ""

        fig_issued, ax_issued = plt.subplots(figsize=(4, 4))
        ax_issued.pie(
            issued_values,
            labels=["Issued By EPC Delivered", "Not Yet Issued"],
            autopct=issued_autopct,
            startangle=140,
            wedgeprops={"width": 0.4}
        )
        ax_issued.set_title("Issued By EPC Delivery", fontsize=9)
        st.pyplot(fig_issued)

    # --------------------------
    # 10) STACKED BAR ISSUED/REVIEW/REPLY BY DISCIPLINE
    # --------------------------
    df["Issued_bool"] = df["Issued By EPC"].notna().astype(int)
    df["Review_bool"] = df["Review By OE"].notna().astype(int)
    df["Reply_bool"] = (df["Reply By EPC"].notna() & (df["Flag"] == 1)).astype(int)

    disc_counts = df.groupby("Discipline")[["Issued_bool", "Review_bool", "Reply_bool"]].sum()

    st.subheader("Number of Docs with Issued, Review, Reply by Discipline")
    fig4, ax4 = plt.subplots(figsize=(8, 5))
    disc_counts.plot(kind="barh", stacked=True, ax=ax4)
    ax4.set_xlabel("Count of Documents", fontsize=9)
    ax4.set_ylabel("Discipline", fontsize=9)
    ax4.set_title("Document Milestone Status by Discipline", fontsize=10)
    plt.xticks(fontsize=8)
    plt.yticks(fontsize=8)
    plt.tight_layout()
    for container in ax4.containers:
        ax4.bar_label(container, label_type='center', fontsize=8)
    st.pyplot(fig4)

    # --------------------------
    # 11) DELAY BY DISCIPLINE (AS OF TODAY)
    # --------------------------
    st.subheader("Delay Percentage by Discipline (As of Today)")

    today = pd.to_datetime("today").normalize()
    actual_prog_today = []
    expected_prog_today = []

    for _, row in df.iterrows():
        a_val = 0.0
        if pd.notna(row["Issued By EPC"]) and row["Issued By EPC"] <= today:
            a_val += ISSUED_WEIGHT
        if pd.notna(row["Review By OE"]) and row["Review By OE"] <= today:
            a_val += REVIEW_WEIGHT
        if pd.not```
/**
 * SPDX-License-Identifier: (MIT OR Apache-2.0)
 */
