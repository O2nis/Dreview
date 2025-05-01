import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime as dt
import numpy as np
import seaborn as sns
import openpyxl

plt.rcParams.update({'font.size': 8})

def parse_date(d):
    """Parse dates, handling Excel serial dates, strings, and empty/invalid entries."""
    if pd.isna(d) or d == '':
        return pd.NaT
    if isinstance(d, (int, float)):  # Handle Excel serial dates
        try:
            return pd.to_datetime(d, unit='D', origin='1899-12-30', errors='coerce')
        except:
            return pd.NaT
    if isinstance(d, str):
        d = d.strip()
        if d in ['0-Jan-00', '########', '']:
            return pd.NaT
        try:
            return pd.to_datetime(d, format='%d-%b-%y', errors='coerce')
        except:
            return pd.to_datetime(d, errors='coerce', dayfirst=True)
    return pd.NaT

def load_data(uploaded_file):
    """Load and preprocess the Excel file."""
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Map expected column names to possible variations
    col_mapping = {
        'Planned Issuance date': ['Planned Issuance date', 'Planned Issuance'],
        'Date Data Upload by EPC': ['Date Data Upload by EPC', 'Actual Upload'],
        'Actual Fichtner Review': ['Actual Fichtner Review', 'Fichtner Review'],
        'Actual EPC reply Date': ['Actual EPC reply Date', 'EPC reply Date'],
        'Final Issuance': ['Final Issuance', 'Final Issuance Date'],
        'Weight': ['Weight', 'Document Weight']
    }
    
    # Find matching columns
    date_cols = []
    for standard_name, possible_names in col_mapping.items():
        found = False
        for name in possible_names:
            if name in df.columns:
                df.rename(columns={name: standard_name}, inplace=True)
                if standard_name != 'Weight':
                    date_cols.append(standard_name)
                found = True
                break
        if not found and standard_name != 'Weight':
            st.warning(f"Could not find column matching: {standard_name}")
    
    # Convert date columns
    for col in date_cols:
        df[col] = df[col].apply(parse_date)
        if df[col].isna().all() and not df[col].empty:
            st.warning(f"No valid dates in column '{col}'. All values: {df[col].unique()}")
    
    # Fill weights
    if 'Weight' in df.columns:
        df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce').fillna(1)
    else:
        df['Weight'] = 1
    
    return df, date_cols

def calculate_curves(df, date_cols, start_date, review_days, final_days, today):
    """Calculate planned and actual S-curves."""
    total_weight = df['Weight'].sum()
    all_dates = set()
    
    # Planned S-curve
    if 'Planned Issuance date' in df.columns and df['Planned Issuance date'].notna().any():
        df['Planned_Review'] = df['Planned Issuance date'] + pd.Timedelta(days=review_days)
        df['Planned_Reply'] = df['Planned_Review'] + pd.Timedelta(days=final_days / 2)
        df['Planned_Final'] = df['Planned_Reply'] + pd.Timedelta(days=final_days / 2)
        
        planned_cols = ['Planned Issuance date', 'Planned_Review', 'Planned_Reply', 'Planned_Final']
        planned_dates = []
        for col in planned_cols:
            valid_dates = df[col].dropna().tolist()
            if not valid_dates:
                st.warning(f"No valid dates in '{col}' for planned S-curve")
            planned_dates.extend(valid_dates)
        all_dates.update(planned_dates)
    else:
        st.error("Missing or empty 'Planned Issuance date' column. Cannot calculate planned S-curve.")
        return pd.DataFrame(), today
    
    # Actual S-curve
    actual_dates = []
    actual_cols = ['Date Data Upload by EPC', ' Bonding Fichtner Review', 'Actual EPC reply Date', 'Final Issuance']
    for col in actual_cols:
        if col in df.columns and df[col].notna().any():
            valid_dates = df[df[col] <= today][col].dropna().tolist()
            if not valid_dates:
                st.warning(f"No valid dates in '{col}' up to today for actual S-curve")
            actual_dates.extend(valid_dates)
        else:
            st.warning(f"Column '{col}' missing or empty for actual S-curve")
    all_dates.update(actual_dates)
    
    if not all_dates:
        st.error("No valid dates found for planned or actual S-curves.")
        return pd.DataFrame(), today
    
    all_dates = sorted(all_dates)
    
    # Calculate progress
    planned_progress = []
    actual_progress = []
    for date in all_dates:
        # Planned progress
        uploaded = df[df['Planned Issuance date'] <= date]['Weight'].sum() if 'Planned Issuance date' in df.columns else 0
        reviewed = df[df['Planned_Review'] <= date]['Weight'].sum() if 'Planned_Review' in df.columns else 0
        replied = df[df['Planned_Reply'] <= date]['Weight'].sum() if 'Planned_Reply' in df.columns else 0
        finalized = df[df['Planned_Final'] <= date]['Weight'].sum() if 'Planned_Final' in df.columns else 0
        planned = (uploaded * 0.25 + reviewed * 0.25 + replied * 0.25 + finalized * 0.25) / total_weight * 100
        planned_progress.append(planned)
        
        # Actual progress
        if date > today:
            actual_progress.append(np.nan)
        else:
            uploaded = df[df['Date Data Upload by EPC'] <= date]['Weight'].sum() if 'Date Data Upload by EPC' in df.columns and df['Date Data Upload by EPC'].notna().any() else 0
            reviewed = df[df['Actual Fichtner Review'] <= date]['Weight'].sum() if 'Actual Fichtner Review' in df.columns and df['Actual Fichtner Review'].notna().any() else 0
            replied = df[df['Actual EPC reply Date'] <= date]['Weight'].sum() if 'Actual EPC reply Date' in df.columns and df['Actual EPC reply Date'].notna().any() else 0
            finalized = df[df['Final Issuance'] <= date]['Weight'].sum() if 'Final Issuance' in df.columns and df['Final Issuance'].notna().any() else 0
            actual = (uploaded * 0.25 + reviewed * 0.25 + replied * 0.25 + finalized * 0.25) / total_weight * 100
            actual_progress.append(actual)
    
    progress_df = pd.DataFrame({
        'Date': all_dates,
        'Planned': planned_progress,
        'Actual': actual_progress
    })
    
    return progress_df, today

def plot_s_curve(progress_df, today, actual_color, expected_color, today_color, end_date_color):
    """Plot the S-curve."""
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # Plot curves
    ax.plot(progress_df['Date'], progress_df['Planned'], label='Planned Progress', color=expected_color, linewidth=2)
    ax.plot(progress_df['Date'], progress_df['Actual'], label='Actual Progress', color=actual_color, linewidth=2)
    
    # Vertical lines
    ax.axvline(today, color=today_color, linestyle='--', linewidth=1.5, label='Today')
    ax.annotate(
        f"Today\n{today.strftime('%d-%b-%Y')}",
        xy=(today, 10),
        xytext=(10, 10),
        textcoords='offset points',
        color=today_color,
        bbox=dict(boxstyle='round,pad=0.3', fc='white', ec='none', alpha=0.7),
        fontsize=8
    )
    
    end_date = progress_df['Date'].max()
    ax.axvline(end_date, color=end_date_color, linestyle='--', linewidth=1.5, label='Expected End')
    ax.annotate(
        f"End\n{end_date.strftime('%d-%b-%Y')}",
        xy=(end_date, 20),
        xytext=(-100, 10),
        textcoords='offset points',
        color=end_date_color,
        bbox=dict(boxstyle='round,pad=0.3', fc='white', ec='none', alpha=0.7),
        fontsize=8,
        arrowprops=dict(arrowstyle='->', color=end_date_color)
    )
    
    # Delay annotation
    if not progress_df.empty:
        last_actual = progress_df[progress_df['Actual'].notna()]['Actual'].iloc[-1]
        planned_today = np.interp(pd.to_numeric(today), pd.to_numeric(progress_df['Date']), progress_df['Planned'])
        delay = planned_today - last_actual
        ax.annotate(
            f"Delay: {delay:.1f}%",
            xy=(today, (last_actual + planned_today) / 2),
            xytext=(10, -50),
            textcoords='offset points',
            color=today_color,
            bbox=dict(boxstyle='round,pad=0.3', fc='white', ec='none', alpha=0.7),
            arrowprops=dict(arrowstyle='->', color=today_color),
            ha='left',
            fontsize=8
        )
    
    # Formatting
    ax.set_title('Document Progress S-Curve', fontsize=12)
    ax.set_xlabel('Date', fontsize=10)
    ax.set_ylabel('Progress (%)', fontsize=10)
    ax.grid(True)
    ax.legend(fontsize=9)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%b-%Y'))
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    return fig

def main():
    st.set_page_config(page_title="Document Register S-Curve Analysis", layout="wide")
    
    # Sidebar inputs
    st.sidebar.header("Configuration")
    uploaded_file = st.sidebar.file_uploader("Upload Document Register Excel", type=['xlsx'])
    start_date = st.sidebar.date_input("Start Date", value=pd.to_datetime('2025-01-01'))
    review_days = st.sidebar.number_input("Review Duration (days)", min_value=1, value=10)
    final_days = st.sidebar.number_input("Final Issuance Duration (days)", min_value=1, value=10)
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Visualization Settings")
    seaborn_style = st.sidebar.selectbox(
        "Select Seaborn Style",
        ["darkgrid", "whitegrid", "dark", "white", "ticks"],
        index=1
    )
    sns.set_style(seaborn_style)
    
    actual_color = st.sidebar.color_picker("Actual Progress Color", "#1f77b4")
    expected_color = st.sidebar.color_picker("Expected Progress Color", "#ff7f0e")
    today_color = st.sidebar.color_picker("Today Line Color", "#000000")
    end_date_color = st.sidebar.color_picker("End Date Line Color", "#d62728")
    
    if uploaded_file is None:
        st.warning("Please upload an Excel file to proceed.")
        return
    
    # Load data
    try:
        df, date_cols = load_data(uploaded_file)
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return
    
    # Display stats
    st.subheader("Document Register Summary")
    st.write(f"Total documents: {len(df)}")
    st.write(f"Total weight: {df['Weight'].sum():.1f}")
    
    if st.checkbox("Show column names"):
        st.write("Columns in uploaded file:", df.columns.tolist())
    
    # Calculate S-curves
    today = pd.to_datetime('2025-05-01')  # Fixed as per instruction
    progress_df, today = calculate_curves(df, date_cols, start_date, review_days, final_days, today)
    
    if not progress_df.empty:
        # Plot S-curve
        st.subheader("Progress S-Curve")
        fig = plot_s_curve(progress_df, today, actual_color, expected_color, today_color, end_date_color)
        st.pyplot(fig)
        
        # Display metrics
        last_actual = progress_df[progress_df['Actual'].notna()]['Actual'].iloc[-1]
        planned_today = np.interp(pd.to_numeric(today), pd.to_numeric(progress_df['Date']), progress_df['Planned'])
        delay = planned_today - last_actual
        st.metric("Current Actual Progress", f"{last_actual:.1f}%")
        st.metric("Planned Progress for Today", f"{planned_today:.1f}%")
        st.metric("Delay", f"{delay:.1f}%", delta_color="inverse")
    else:
        st.error("Unable to generate S-curve due to missing or invalid data.")

if __name__ == "__main__":
    main()
