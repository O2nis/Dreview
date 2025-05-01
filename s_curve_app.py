import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime as dt
import numpy as np
from io import BytesIO
import openpyxl

def parse_date(d):
    """Robust date parsing for document register format, including Excel serial dates."""
    if pd.isna(d) or d in ['', '0-Jan-00', '########']:
        return pd.NaT
    if isinstance(d, (int, float)):  # Handle Excel serial dates
        try:
            return pd.to_datetime(d, unit='D', origin='1899-12-30', errors='coerce')
        except:
            return pd.NaT
    try:
        return pd.to_datetime(d, format='%d-%b-%y', errors='coerce')
    except:
        return pd.to_datetime(d, errors='coerce', dayfirst=True)

def calculate_document_progress(df, today):
    """
    Calculate progress percentages for each document based on milestones.
    Returns: DataFrame with progress columns added.
    """
    # Calculate progress at each milestone (0% if NaT or future date)
    df['progress_upload'] = (df['Date Data Upload by EPC'].notna() & 
                           (df['Date Data Upload by EPC'] <= today)).astype(int) * 0.25
    df['progress_review'] = (df['Actual Fichtner Review'].notna() & 
                           (df['Actual Fichtner Review'] <= today)).astype(int) * 0.25
    df['progress_reply'] = (df['Actual EPC reply Date'].notna() & 
                          (df['Actual EPC reply Date'] <= today)).astype(int) * 0.25
    df['progress_final'] = (df['Final Issuance'].notna() & 
                          (df['Final Issuance'] <= today)).astype(int) * 0.25
    
    # Calculate total actual progress
    df['actual_progress'] = (df['progress_upload'] + df['progress_review'] + 
                           df['progress_reply'] + df['progress_final']) * df['Weight']
    
    return df

def generate_s_curve(df, start_date, review_days=10, final_days=10):
    """Generate S-curve data for planned vs actual progress."""
    today = pd.to_datetime('2025-05-01')  # Fixed as per instruction
    
    # Calculate planned dates
    if 'Planned Issuance date' in df.columns and df['Planned Issuance date'].notna().any():
        df['Planned_Review'] = df['Planned Issuance date'] + pd.Timedelta(days=review_days)
        df['Planned_Reply'] = df['Planned_Review'] + pd.Timedelta(days=final_days / 2)
        df['Planned_Final'] = df['Planned_Reply'] + pd.Timedelta(days=final_days / 2)
    else:
        st.error("Missing or empty 'Planned Issuance date' column.")
        return None, None, today
    
    # Get all relevant dates
    date_columns = [
        'Planned Issuance date', 'Planned_Review', 'Planned_Reply', 'Planned_Final',
        'Date Data Upload by EPC', 'Actual Fichtner Review', 
        'Actual EPC reply Date', 'Final Issuance'
    ]
    
    all_dates = set()
    for col in date_columns:
        if col in df.columns:
            valid_dates = df[col].dropna()
            if valid_dates.empty:
                st.warning(f"No valid dates in '{col}'")
            all_dates.update(valid_dates)
    
    if not all_dates:
        st.error("No valid dates found for S-curve.")
        return None, None, today
    
    all_dates = sorted(all_dates)
    total_weight = df['Weight'].sum()
    
    # Calculate cumulative progress
    planned_progress = []
    actual_progress = []
    
    for date in all_dates:
        # Planned progress
        planned_upload = df[df['Planned Issuance date'] <= date]['Weight'].sum() if 'Planned Issuance date' in df.columns else 0
        planned_review = df[df['Planned_Review'] <= date]['Weight'].sum() if 'Planned_Review' in df.columns else 0
        planned_reply = df[df['Planned_Reply'] <= date]['Weight'].sum() if 'Planned_Reply' in df.columns else 0
        planned_final = df[df['Planned_Final'] <= date]['Weight'].sum() if 'Planned_Final' in df.columns else 0
        planned_value = (planned_upload * 0.25 + planned_review * 0.25 + 
                        planned_reply * 0.25 + planned_final * 0.25) / total_weight * 100
        planned_progress.append(planned_value)
        
        # Actual progress
        if date > today:
            actual_progress.append(np.nan)
        else:
            actual_upload = df[df['Date Data Upload by EPC'] <= date]['Weight'].sum() if 'Date Data Upload by EPC' in df.columns else 0
            actual_review = df[df['Actual Fichtner Review'] <= date]['Weight'].sum() if 'Actual Fichtner Review' in df.columns else 0
            actual_reply = df[df['Actual EPC reply Date'] <= date]['Weight'].sum() if 'Actual EPC reply Date' in df.columns else 0
            actual_final = df[df['Final Issuance'] <= date]['Weight'].sum() if 'Final Issuance' in df.columns else 0
            actual_value = (actual_upload * 0.25 + actual_review * 0.25 + 
                           actual_reply * 0.25 + actual_final * 0.25) / total_weight * 100
            actual_progress.append(actual_value)
    
    progress_df = pd.DataFrame({
        'Date': all_dates,
        'Planned': planned_progress,
        'Actual': actual_progress
    }).sort_values('Date')
    
    return progress_df, total_weight, today

def plot_s_curve(progress_df, total_weight, today):
    """Create the S-curve visualization."""
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Plot curves
    ax.plot(progress_df['Date'], progress_df['Planned'], label='Planned', color='#1f77b4', linewidth=2)
    ax.plot(progress_df['Date'], progress_df['Actual'], label='Actual', color='#ff7f0e', linewidth=2)
    
    # Add today line
    ax.axvline(today, color='#d62728', linestyle='--', linewidth=1.5, label='Today')
    
    # Add end date line
    end_date = progress_df['Date'].max()
    ax.axvline(end_date, color='#2ca02c', linestyle='--', linewidth=1.5, label='Expected End')
    
    # Calculate delay
    last_actual = progress_df[progress_df['Actual'].notna()]['Actual'].iloc[-1]
    # Convert dates to numeric for interpolation
    date_numeric = pd.to_numeric(progress_df['Date'].astype('datetime64[ns]'))
    today_numeric = pd.to_numeric(today.astype('datetime64[ns]'))
    planned_today = np.interp(today_numeric, date_numeric, progress_df['Planned'])
    delay = planned_today - last_actual
    
    # Add delay annotation
    ax.annotate(
        f"Delay: {delay:.1f}%",
        xy=(today, (last_actual + planned_today)/2),
        xytext=(10, 0),
        textcoords="offset points",
        color='#d62728',
        bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.7),
        arrowprops=dict(arrowstyle="->", color='#d62728')
    )
    
    # Formatting
    ax.set_title(f'Document Progress S-Curve (Total Weight: {total_weight:.1f})')
    ax.set_xlabel('Date')
    ax.set_ylabel('Progress (%)')
    ax.legend()
    ax.grid(True)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%b-%Y'))
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    return fig

def main():
    st.set_page_config(page_title="Document S-Curve Analysis", layout="wide")
    st.title("📊 Document Register S-Curve Analysis")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Document Register Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is None:
        st.info("Please upload your document register file")
        return
    
    # Load and process data
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Standardize column names
        df.columns = df.columns.str.strip()
        
        # Map column names
        col_mapping = {
            'Planned Issuance date': ['Planned Issuance date', 'Planned Issuance'],
            'Date Data Upload by EPC': ['Date Data Upload by EPC', 'Actual Upload'],
            'Actual Fichtner Review': ['Actual Fichtner Review', 'Fichtner Review'],
            'Actual EPC reply Date': ['Actual EPC reply Date', 'EPC reply Date'],
            'Final Issuance': ['Final Issuance', 'Final Issuance Date'],
            'Weight': ['Weight', 'Document Weight']
        }
        
        for standard_name, possible_names in col_mapping.items():
            for name in possible_names:
                if name in df.columns:
                    df.rename(columns={name: standard_name}, inplace=True)
                    break
        
        # Convert date columns
        date_cols = [
            'Planned Issuance date',
            'Date Data Upload by EPC',
            'Actual Fichtner Review',
            'Actual EPC reply Date',
            'Final Issuance'
        ]
        
        for col in date_cols:
            if col in df.columns:
                df[col] = df[col].apply(parse_date)
                if df[col].isna().all() and not df[col].empty:
                    st.warning(f"No valid dates in '{col}'. All values: {df[col].unique()}")
        
        # Handle weights
        if 'Weight' not in df.columns:
            df['Weight'] = 1
        else:
            df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce').fillna(1)
        
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return
    
    # Configuration
    st.sidebar.header("Settings")
    start_date = st.sidebar.date_input("Start Date", value=pd.to_datetime('2025-01-01'))
    review_days = st.sidebar.number_input("Review Duration (days)", min_value=1, value=10)
    final_days = st.sidebar.number_input("Final Issuance Duration (days)", min_value=1, value=10)
    
    # Calculate progress
    df = calculate_document_progress(df, today=pd.to_datetime('2025-05-01'))
    progress_df, total_weight, today = generate_s_curve(df, start_date, review_days, final_days)
    
    if progress_df is None:
        st.error("Could not generate S-curve - check your date columns")
        return
    
    # Display S-curve
    st.subheader("Document Progress S-Curve")
    fig = plot_s_curve(progress_df, total_weight, today)
    st.pyplot(fig)
    
    # Display current status
    last_actual = progress_df[progress_df['Actual'].notna()]['Actual'].iloc[-1]
    date_numeric = pd.to_numeric(progress_df['Date'].astype('datetime64[ns]'))
    today_numeric = pd.to_numeric(today.astype('datetime64[ns]'))
    planned_today = np.interp(today_numeric, date_numeric, progress_df['Planned'])
    delay = planned_today - last_actual
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Actual Progress", f"{last_actual:.1f}%")
    col2.metric("Planned Progress", f"{planned_today:.1f}%")
    col3.metric("Delay", f"{delay:.1f}%", delta_color="inverse")
    
    # Show data table
    if st.checkbox("Show progress data"):
        st.dataframe(progress_df)

if __name__ == "__main__":
    main()
