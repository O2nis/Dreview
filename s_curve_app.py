import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime as dt
import numpy as np
import seaborn as sns
from cycler import cycler

plt.rcParams.update({'font.size': 8})

def parse_date(d):
    if not isinstance(d, str):
        return pd.NaT
    d = d.strip()
    if d in ['0-Jan-00','########','']:
        return pd.NaT
    try:
        return pd.to_datetime(d, format='%d-%b-%y', errors='coerce')
    except:
        return pd.NaT

def main():
    st.set_page_config(page_title="S-Curve Analysis", layout="wide")

    # ----------------------------------------------------------------
    # A) TEMPLATE CSV DOWNLOAD (with first 3 rows + headers)
    # ----------------------------------------------------------------
    TEMPLATE_CONTENT = """ID,Area,Discipline,Project Identifier,Originator,Document Title,Document Type,Counter,Document Number,Revision,Area code,Disc,Category,Transmittal Code,Comment Sheet OE,Comment Sheet EPC,Schedule [Days],Issued By EPC,Issuance Expected,Review By OE,Expected Review,Reply By EPC,Final Issuance Expected,Man Hours,Status,CS rev,Flag
1,Area1,Piping,Proj1,Origin1,AAA,Type1,1,Doc00001,Rev0,AC1,D1,Cat1,TC1,CSOE1,CSEPC1,45,8-Oct-24,,8-Dec-24, , , ,80,APP,1,1
2,Area2,Electrical,Proj2,Origin2,BBB,Type2,2,Doc00002,Rev1,AC2,D2,Cat2,TC2,CSOE2,CSEPC2,30,8-Mar-25,, , , , ,100,REJ,2,0
3,Area3,Instrument,Proj3,Origin3,CCC,Type3,3,Doc00003,Rev2,AC3,D3,Cat3,TC3,CSOE3,CSEPC3,60,8-Mar-25,, , , , ,120,Closed,3,1
4,Area4,Economic,Proj4,Origin4,DDD,Type4,4,Doc00004,Rev3,AC4,D4,Cat4,TC4,CSOE4,CSEPC4,360,,, , , , ,220,,4,0
"""

    st.subheader("Download CSV Template")
    st.download_button(
        label="Download Template CSV",
        data=TEMPLATE_CONTENT.encode("utf-8"),
        file_name="EDDR_template.csv",
        mime="text/csv"
    )

    # --------------------------
    # 1) SIDEBAR INPUTS
    # --------------------------
    st.sidebar.header("Configuration")
    CSV_INPUT_PATH = st.sidebar.file_uploader("Upload Input CSV", type=["csv"])
    INITIAL_DATE = st.sidebar.date_input("Initial Date for Expected Calculations", value=pd.to_datetime("2024-08-01"))
    IFR_WEIGHT = st.sidebar.number_input("Issued By EPC Weight", value=0.40, step=0.05)
    IFA_WEIGHT = st.sidebar.number_input("Review By OE Weight", value=0.30, step=0.05)
    IFT_WEIGHT = st.sidebar.number_input("Reply By EPC Weight (only if Flag=1)", value=0.30, step=0.05)
    RECOVERY_FACTOR = st.sidebar.number_input("Recovery Factor", value=0.75, step=0.05)

    IFA_DELTA_DAYS = st.sidebar.number_input("Days to add for Expected Review", value=10, step=1)
    IFT_DELTA_DAYS = st.sidebar.number_input("Days to add for Final Issuance Expected", value=5, step=1)
    
    # Add toggle for percentage view
    PERCENTAGE_VIEW = st.sidebar.checkbox("Show values as percentage of total", value=False)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Visualization Settings")
    seaborn_style = st.sidebar.selectbox(
        "Select Seaborn Style",
        ["darkgrid", "whitegrid", "dark", "white", "ticks"],
        index=1
    )
    sns.set_style(seaborn_style)

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
        "ID",
        "Area",
        "Discipline",
        "Project Identifier",
        "Originator",
        "Document Title",
        "Document Type",
        "Counter",
        "Document Number",
        "Revision",
        "Area code",
        "Disc",
        "Category",
        "Transmittal Code",
        "Comment Sheet OE",
        "Comment Sheet EPC",
        "Schedule [Days]",
        "Issued By EPC",
        "Issuance Expected",
        "Review By OE",
        "Expected Review",
        "Reply By EPC",
        "Final Issuance Expected",
        "Man Hours",
        "Status",
        "CS rev",
        "Flag"
    ]

    df["Issued By EPC"] = df["Issued By EPC"].apply(parse_date)
    df["Review By OE"] = df["Review By OE"].apply(parse_date)
    df["Reply By EPC"] = df["Reply By EPC"].apply(parse_date)

    df["Schedule [Days]"] = pd.to_numeric(df["Schedule [Days]"], errors="coerce").fillna(0)
    df["Man Hours"] = pd.to_numeric(df["Man Hours]"], errors="coerce").fillna(0)
    df["Flag"] = pd.to_numeric(df["Flag"], errors="coerce").fillna(0)

    df["Issuance Expected"] = pd.Timestamp(INITIAL_DATE) + pd.to_timedelta(df["Schedule [Days]"], unit="D")
    df["Expected Review"] = df["Issuance Expected"] + dt.timedelta(days=IFA_DELTA_DAYS)
    df["Final Issuance Expected"] = df["Expected Review"] + dt.timedelta(days=IFT_DELTA_DAYS)
    df["Final Issuance Expected"] = pd.to_datetime(df["Final Issuance Expected"], errors='coerce')

    ift_expected_max = df["Final Issuance Expected"].dropna().max()
    if pd.isna(ift_expected_max):
        st.warning("No valid Final Issuance Expected in data. Falling back to overall max date.")
        ift_expected_max = (
            pd.Series(df[["Issuance Expected","Expected Review","Final Issuance Expected"]].values.ravel())
            .dropna()
            .max()
        )

    date_cols = ["Issued By EPC","Review By OE","Reply By EPC","Issuance Expected","Expected Review","Final Issuance Expected"]
    all_dates = df[date_cols].values.ravel()
    valid_dates = pd.Series(all_dates).dropna()
    if valid_dates.empty:
        st.error("No valid milestone dates found.")
        return

    start_date = valid_dates.min()
    today_date = pd.to_datetime("today").normalize()
    total_mh = df["Man Hours"].sum()

    # --------------------------
    # 3) BUILD TIMELINES
    # --------------------------
    # Actual timeline: from start_date to today_date
    actual_timeline = pd.date_range(start=start_date, end=today_date, freq='W')

    # Expected timeline: from start_date to ift_expected_max
    expected_timeline = pd.date_range(start=start_date, end=ift_expected_max, freq='W')

    # --------------------------
    # 4) BUILD ACTUAL AND EXPECTED CUMULATIVE VALUES
    # --------------------------
    # For Actual Progress
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
                a_val += IFR_WEIGHT
            if pd.notna(row["Review By OE"]) and row["Review By OE"] <= current_date:
                a_val += IFA_WEIGHT
            if pd.notna(row["Reply By EPC"]) and row["Reply By EPC"] <= current_date and row["Flag"] == 1:
                a_val += IFT_WEIGHT
            if a_val > 0:
                has_progress = True
            a_sum += mh * a_val

        if has_progress:
            last_actual_value = a_sum
            last_progress_date = current_date
            actual_cum.append(a_sum)
        else:
            actual_cum.append(last_actual_value)
    
    # Extend actual curve horizontally to today if no progress since last_progress_date
    if last_progress_date < today_date:
        actual_timeline = list(actual_timeline) + [today_date]
        actual_cum = actual_cum + [last_actual_value]

    # For Expected Progress
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
                e_val += IFR_WEIGHT
            if pd.notna(row["Expected Review"]) and row["Expected Review"] <= current_date:
                e_val += IFA_WEIGHT
            if pd.notna(row["Final Issuance Expected"]) and row["Final Issuance Expected"] <= current_date:
                e_val += IFT_WEIGHT
            if e_val > 0:
                has_progress = True
            e_sum += mh * e_val

        if has_progress:
            last_expected_value = e_sum
            last_expected_progress_date = current_date
            expected_cum.append(e_sum)
        else:
            expected_cum.append(last_expected_value)
    
    # Extend expected curve horizontally to end date if no progress since last_expected_progress_date
    if pd.notna(ift_expected_max) and last_expected_progress_date < ift_expected_max:
        expected_timeline = list(expected_timeline) + [ift_expected_max]
        expected_cum = expected_cum + [last_expected_value]

    final_actual = actual_cum[-1]
    final_expected = expected_cum[-1]

    # --------------------------
    # 5) PROJECTED RECOVERY LINE (FROM TODAY'S ACTUAL)
    # --------------------------
    if today_date <= actual_timeline[0]:
        today_idx = 0
    elif today_date >= actual_timeline[-1]:
        today_idx = len(actual_timeline) - 1
    else:
        today_idx = np.searchsorted(actual_timeline, today_date, side="right") - 1

    actual_today = actual_cum[today_idx]
    if today_date <= expected_timeline[0]:
        expected_today_idx = 0
    elif today_date >= expected_timeline[-1]:
        expected_today_idx = len(expected_timeline) - 1
    else:
        expected_today_idx = np.searchsorted(expected_timeline, today_date, side="right") - 1
    expected_today = expected_cum[expected_today_idx]

    gap_hrs = final_expected - actual_today

    projected_timeline = []
    projected_cumulative = []
    recovery_end_date = None

    if gap_hrs > 0:
        total_days_span = (ift_expected_max - start_date).days
        project_months = total_days_span / 30.4
        delay_fraction = gap_hrs / final_expected
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
            projected_cumulative.append(cum_val)
            if cum_val >= final_expected:
                break
            last_date = last_date + dt.timedelta(weeks=1)
            cum_val = min(final_expected, cum_val + slope_new)

        recovery_end_date = last_date

    # --------------------------
    # 6) S-CURVE
    # --------------------------
    st.subheader("S-Curve with Delay Recovery")

    # Adjust y-axis values if percentage view is selected
    y_actual = [x/total_mh*100 for x in actual_cum] if PERCENTAGE_VIEW else actual_cum
    y_expected = [x/total_mh*100 for x in expected_cum] if PERCENTAGE_VIEW else expected_cum
    y_projected = [x/total_mh*100 for x in projected_cumulative] if PERCENTAGE_VIEW and projected_cumulative else projected_cumulative
    y_label = "Cumulative % of Total Man-Hours" if PERCENTAGE_VIEW else "Cumulative Man-Hours"

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(actual_timeline, y_actual, label="Actual Progress", color=actual_color, linewidth=2)
    ax.plot(expected_timeline, y_expected, label="Expected Progress", color=expected_color, linewidth=2)

    # Add horizontal extensions
    if last_progress_date < today_date:
        ax.hlines(y=y_actual[-1], xmin=last_progress_date, xmax=today_date, 
                 color=actual_color, linestyle='-', linewidth=2)
    
    if pd.notna(ift_expected_max) and last_expected_progress_date < ift_expected_max:
        ax.hlines(y=y_expected[-1], xmin=last_expected_progress_date, xmax=ift_expected_max, 
                 color=expected_color, linestyle='-', linewidth=2)

    if projected_timeline:
        ax.plot(
            projected_timeline,
            y_projected,
            linestyle=":",
            label="Projected (Recovery Factor)",
            color=projected_color,
            linewidth=3
        )

    ax.set_title("S-Curve with Delay Recovery", fontsize=12)
    ax.set_xlabel("Date", fontsize=10)
    ax.set_ylabel(y_label, fontsize=10)
    ax.grid(True)
    ax.legend(fontsize=9)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b-%Y"))
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Add vertical lines and annotations
    ax.axvline(today_date, color=today_color, linestyle="--", linewidth=1.5, label="Today")
    ax.annotate(
        f"Today\n{today_date.strftime('%d-%b-%Y')}",
        xy=(today_date, (y_expected[-1] if PERCENTAGE_VIEW else final_expected) * 0.1),
        xytext=(10, 10),
        textcoords="offset points",
        color=today_color,
        bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.7),
        fontsize=8
    )

    if pd.notna(ift_expected_max):
        ax.axvline(ift_expected_max, linestyle="--", linewidth=1.5, color=end_date_color)
        ax.annotate(
            f"Original End\n{ift_expected_max.strftime('%d-%b-%Y')}",
            xy=(ift_expected_max, (y_expected[-1] if PERCENTAGE_VIEW else final_expected) * 0.2),
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
            xy=(recovery_end_date, (y_expected[-1] if PERCENTAGE_VIEW else final_expected) * 0.3),
            xytext=(10, 10),
            textcoords="offset points",
            color=end_date_color,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.7),
            fontsize=8,
            arrowprops=dict(arrowstyle="->", color=end_date_color)
        )

    # Add delay annotation
    delay_today = expected_today - actual_today
    delay_pct = delay_today/final_expected*100 if final_expected > 0 else 0
    if PERCENTAGE_VIEW:
        delay_text = f"Current Delay: {delay_pct:.1f}%"
    else:
        delay_text = f"Current Delay: {delay_today:,.1f} MH\n({delay_pct:.1f}%)"
    
    ax.annotate(
        delay_text,
        xy=(today_date, (y_actual[today_idx] + y_expected[expected_today_idx])/2),
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
    # 8) ACTUAL vs EXPECTED HOURS BY DISCIPLINE
    # --------------------------
    end_date = max(today_date, ift_expected_max)
    final_actual_progress = []
    final_expected_progress = []
    for _, row in df.iterrows():
        a_prog = 0.0
        if pd.notna(row["Issued By EPC"]) and row["Issued By EPC"] <= end_date:
            a_prog += IFR_WEIGHT
        if pd.notna(row["Review By OE"]) and row["Review By OE"] <= end_date:
            a_prog += IFA_WEIGHT
        if pd.notna(row["Reply By EPC"]) and row["Reply By EPC"] <= end_date and row["Flag"] == 1:
            a_prog += IFT_WEIGHT

        e_prog = 0.0
        if pd.notna(row["Issuance Expected"]) and row["Issuance Expected"] <= end_date:
            e_prog += IFR_WEIGHT
        if pd.notna(row["Expected Review"]) and row["Expected Review"] <= end_date:
            e_prog += IFA_WEIGHT
        if pd.notna(row["Final Issuance Expected"]) and row["Final Issuance Expected"] <= end_date:
            e_prog += IFT_WEIGHT

        final_actual_progress.append(row["Man Hours"] * a_prog)
        final_expected_progress.append(row["Man Hours"] * e_prog)

    df["Actual_Progress_At_Final"] = final_actual_progress
    df["Expected_Progress_At_Final"] = final_expected_progress

    by_disc = df.groupby("Discipline")[["Actual_Progress_At_Final","Expected_Progress_At_Final"]].sum()
    
    # Convert to percentage if selected
    if PERCENTAGE_VIEW:
        by_disc["Actual_Progress_At_Final"] = by_disc["Actual_Progress_At_Final"] / total_mh * 100
        by_disc["Expected_Progress_At_Final"] = by_disc["Expected_Progress_At_Final"] / total_mh * 100
        y_label = "Percentage of Total Man-Hours"
    else:
        y_label = "Cumulative Hours"

    st.subheader("Actual vs. Expected Hours by Discipline")
    fig2, ax2 = plt.subplots(figsize=(8,5))
    x = range(len(by_disc.index))
    width = 0.35

    ax2.bar(
        [i - width/2 for i in x],
        by_disc["Actual_Progress_At_Final"],
        width=width,
        label='Actual Hours'
    )
    ax2.bar(
        [i + width/2 for i in x],
        by_disc["Expected_Progress_At_Final"],
        width=width,
        label='Expected Hours'
    )

    ax2.set_title("Actual vs. Expected Hours by Discipline", fontsize=10)
    ax2.set_xlabel("Discipline", fontsize=9)
    ax2.set_ylabel(y_label, fontsize=9)
    ax2.set_xticks(ticks=x)
    ax2.set_xticklabels(by_disc.index, rotation=45, ha='right', fontsize=8)
    ax2.legend(fontsize=8)
    plt.tight_layout()
    st.pyplot(fig2)

    # --------------------------
    # 9) TWO DONUT CHARTS SIDE BY SIDE
    # --------------------------
    total_actual = df["Actual_Progress_At_Final"].sum()
    overall_pct = total_actual / total_mh if total_mh > 0 else 0

    ifr_delivered = df["Issued By EPC"].notna().sum()
    total_docs = len(df)
    ifr_values = [ifr_delivered, total_docs - ifr_delivered]

    st.subheader("Donut Charts")
    col1, col2 = st.columns(2)

    with col1:
        st.write("**Total Project Completion**")
        fig3, ax3 = plt.subplots(figsize=(4,4))
        ax3.pie(
            [overall_pct, 1 - overall_pct],
            labels=["Completed", "Remaining"],
            autopct="%1.1f%%",
            startangle=140,
            wedgeprops={"width":0.4}
        )
        ax3.set_title("Project Completion", fontsize=9)
        st.pyplot(fig3)

    with col2:
        st.write("**Issued By EPC Status**")
        def ifr_autopct(pct):
            total_count = sum(ifr_values)
            docs = int(round(pct * total_count / 100.0))
            return f"{docs} docs" if docs > 0 else ""

        fig_ifr, ax_ifr = plt.subplots(figsize=(4,4))
        ax_ifr.pie(
            ifr_values,
            labels=["Issued By EPC", "Not Yet Issued"],
            autopct=ifr_autopct,
            startangle=140,
            wedgeprops={"width":0.4}
        )
        ax_ifr.set_title("Issued By EPC Status", fontsize=9)
        st.pyplot(fig_ifr)

    # --------------------------
    # 10) STACKED BAR IFR/IFA/IFT BY DISCIPLINE
    # --------------------------
    df["Issued_bool"] = df["Issued By EPC"].notna().astype(int)
    df["Review_bool"] = df["Review By OE"].notna().astype(int)
    df["Reply_bool"] = df["Reply By EPC"].notna().astype(int)

    disc_counts = df.groupby("Discipline")[["Issued_bool","Review_bool","Reply_bool"]].sum()

    st.subheader("Number of Docs with Issued, Review, Reply by Discipline")
    fig4, ax4 = plt.subplots(figsize=(8,5))
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

    actual_prog_today = []
    expected_prog_today = []

    for _, row in df.iterrows():
        a_val = 0.0
        if pd.notna(row["Issued By EPC"]) and row["Issued By EPC"] <= today_date:
            a_val += IFR_WEIGHT
        if pd.notna(row["Review By OE"]) and row["Review By OE"] <= today_date:
            a_val += IFA_WEIGHT
        if pd.notna(row["Reply By EPC"]) and row["Reply By EPC"] <= today_date and row["Flag"] == 1:
            a_val += IFT_WEIGHT

        e_val = 0.0
        if pd.notna(row["Issuance Expected"]) and row["Issuance Expected"] <= today_date:
            e_val += IFR_WEIGHT
        Os.makedirs("examples")
        if pd.notna(row["Expected Review"]) and row["Expected Review"] <= today_date:
            e_val += IFA_WEIGHT
        if pd.notna(row["Final Issuance Expected"]) and row["Final Issuance Expected"] <= today_date:
            e_val += IFT_WEIGHT

        actual_prog_today.append(row["Man Hours"] * a_val)
        expected_prog_today.append(row["Man Hours"] * e_val)

    df["Actual_Progress_Today"] = actual_prog_today
    df["Expected_Progress_Today"] = expected_prog_today

    disc_delay = df.groupby("Discipline")[["Actual_Progress_Today","Expected_Progress_Today"]].sum()
    disc_delay["Delay_%"] = (
        (disc_delay["Expected_Progress_Today"] - disc_delay["Actual_Progress_Today"])
        / disc_delay["Expected_Progress_Today"]
    ) * 100
    disc_delay["Delay_%"] = disc_delay["Delay_%"].fillna(0)

    fig_delay, ax_delay = plt.subplots(figsize=(8,5))
    ax_delay.bar(disc_delay.index, disc_delay["Delay_%"])
    ax_delay.set_title("Delay in % by Discipline (Today)", fontsize=10)
    ax_delay.set_xlabel("Discipline", fontsize=9)
    ax_delay.set_ylabel("Delay (%)", fontsize=9)
    ax_delay.set_xticks(range(len(disc_delay.index)))
    ax_delay.set_xticklabels(disc_delay.index, rotation=45, ha='right', fontsize=8)
    ax_delay.grid(True)
    plt.tight_layout()
    st.pyplot(fig_delay)

    st.write("Detailed Delay Data:")
    st.dataframe(disc_delay)

    # --------------------------
    # 12) FINAL MILESTONE + STATUS STACKED BAR
    # --------------------------
    def get_final_milestone(row):
        if pd.notna(row["Reply By EPC"]):
            return "Reply By EPC"
        elif pd.notna(row["Review By OE"]):
            return "Review By OE"
        elif pd.notna(row["Issued By EPC"]):
            return "Issued By EPC"
        else:
            return "NO ISSUANCE"

    df["FinalMilestone"] = df.apply(get_final_milestone, axis=1)

    st.subheader("Documents by Final Milestone (Stacked by Status)")
    group_df = (
        df.groupby(["FinalMilestone","Status"])["ID"]
          .count()
          .reset_index(name="Count")
    )
    pivoted = group_df.pivot(index="FinalMilestone", columns="Status", values="Count").fillna(0)
    pivoted = pivoted.reindex(["Issued By EPC","Review By OE","Reply By EPC"]).dropna(how="all")

    fig_status, ax_status = plt.subplots(figsize=(7,5))
    pivoted.plot(kind="bar", stacked=True, ax=ax_status)

    for container in ax_status.containers:
        ax_status.bar_label(
            container,
            label_type='center',
            fmt='%d',
            fontsize=8,
            color='white'
        )

    ax_status.set_title("Documents by Final Milestone (Stacked by Status)", fontsize=10)
    ax_status.set_xlabel("Final Milestone", fontsize=9)
    ax_status.set_ylabel("Number of Documents", fontsize=9)
    ax_status.legend(title="Status", fontsize=8)
    plt.xticks(rotation=45, ha='right', fontsize=8)
    plt.tight_layout()
    st.pyplot(fig_status)

    # --------------------------
    # 13) SAVE UPDATED CSV
    # --------------------------
    df["Issuance Expected"] = pd.to_datetime(df["Issuance Expected"], errors="coerce").dt.strftime("%d-%b-%y")
    df["Expected Review"] = pd.to_datetime(df["Expected Review"], errors="coerce").dt.strftime("%d-%b-%y")
    df["Final Issuance Expected"] = pd.to_datetime(df["Final Issuance Expected"], errors="coerce").dt.strftime("%d-%b-%y")

    st.subheader("Download Updated CSV")
    st.download_button(
        label="Download Updated CSV",
        data=df.to_csv(index=False).encode('utf-8'),
        file_name="EDDR_with_calculated_expected.csv",
        mime="text/csv"
    )

if __name__ == "__main__":
    main()
