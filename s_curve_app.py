import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, timedelta
from io import BytesIO

def load_data(uploaded_file):
    """Load and preprocess the Excel file with robust error handling"""
    try:
        # Read Excel file into DataFrame
        df = pd.read_excel(uploaded_file)
        
        # Clean column names by stripping whitespace and standardizing
        df.columns = df.columns.str.strip().str.replace(' ', '_').str.lower()
        
        # Create column name mapping (common variations)
        column_map = {
            'planned_issuance_date': ['planned_issuance_date', 'planned_date', 'planned'],
            'date_data_upload_by_epc': ['date_data_upload_by_epc', 'upload_date', 'actual_upload'],
            'actual_fichtner_review': ['actual_fichtner_review', 'review_date', 'fichtner_review'],
            'actual_epc_reply_date': ['actual_epc_reply_date', 'reply_date', 'epc_reply'],
            'final_issuance': ['final_issuance', 'final_date', 'completion_date'],
            'weight': ['weight', 'document_weight', 'importance']
        }
        
        # Standardize column names
        for standard_name, alternatives in column_map.items():
            for alt in alternatives:
                if alt in df.columns:
                    df.rename(columns={alt: standard_name}, inplace=True)
                    break
        
        # Convert date columns - handle errors gracefully
        date_columns = [
            'planned_issuance_date',
            'date_data_upload_by_epc',
            'actual_fichtner_review',
            'actual_epc_reply_date',
            'final_issuance'
        ]
        
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Handle weights - default to 1 if not specified
        if 'weight' not in df.columns:
            df['weight'] = 1
        else:
            df['weight'] = pd.to_numeric(df['weight'], errors='coerce').fillna(1)
        
        return df
    
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def calculate_progress(df, review_days=10, final_days=10):
    """Calculate progress curves with robust date handling"""
    if df is None or len(df) == 0:
        return None, None, datetime.now()
    
    today = datetime.now()
    
    # Calculate planned milestones
    if 'planned_issuance_date' in df.columns:
        df['planned_review'] = df['planned_issuance_date'] + timedelta(days=review_days)
        df['planned_final'] = df['planned_review'] + timedelta(days=final_days)
    
    # Collect all relevant dates
    all_dates = set()
    
    # Planned dates
    planned_phases = ['planned_issuance_date', 'planned_review', 'planned_final']
    for phase in planned_phases:
        if phase in df.columns:
            valid_dates = df[phase].dropna()
            if not valid_dates.empty:
                all_dates.update(valid_dates)
    
    # Actual dates
    actual_phases = [
        'date_data_upload_by_epc',
        'actual_fichtner_review',
        'actual_epc_reply_date',
        'final_issuance'
    ]
    for phase in actual_phases:
        if phase in df.columns:
            valid_dates = df[phase].dropna()
            if not valid_dates.empty:
                all_dates.update(valid_dates)
    
    if not all_dates:
        return None, None, today
    
    # Convert to sorted list of dates
    all_dates = sorted([d for d in all_dates if isinstance(d, datetime)])
    
    # Calculate progress over time
    planned_progress = []
    actual_progress = []
    total_weight = df['weight'].sum()
    
    for date in all_dates:
        # Planned progress calculation
        planned_upload = planned_review = planned_final = 0
        if 'planned_issuance_date' in df.columns:
            planned_upload = df[df['planned_issuance_date'] <= date]['weight'].sum()
        if 'planned_review' in df.columns:
            planned_review = df[df['planned_review'] <= date]['weight'].sum()
        if 'planned_final' in df.columns:
            planned_final = df[df['planned_final'] <= date]['weight'].sum()
        
        planned_value = (planned_upload * 0.25 + planned_review * 0.25 + planned_final * 0.5) / total_weight * 100
        planned_progress.append(planned_value)
        
        # Actual progress calculation (only up to today)
        actual_upload = actual_review = actual_reply = actual_final = 0
        if date <= today:
            if 'date_data_upload_by_epc' in df.columns:
                actual_upload = df[df['date_data_upload_by_epc'] <= date]['weight'].sum()
            if 'actual_fichtner_review' in df.columns:
                actual_review = df[df['actual_fichtner_review'] <= date]['weight'].sum()
            if 'actual_epc_reply_date' in df.columns:
                actual_reply = df[df['actual_epc_reply_date'] <= date]['weight'].sum()
            if 'final_issuance' in df.columns:
                actual_final = df[df['final_issuance'] <= date]['weight'].sum()
            
            actual_value = (actual_upload * 0.25 + actual_review * 0.25 + 
                           actual_reply * 0.25 + actual_final * 0.25) / total_weight * 100
        else:
            actual_value = np.nan
        
        actual_progress.append(actual_value)
    
    # Create progress DataFrame
    progress_df = pd.DataFrame({
        'date': all_dates,
        'planned': planned_progress,
        'actual': actual_progress
    }).sort_values('date')
    
    return progress_df, total_weight, today

def plot_s_curve(progress_df, total_weight, today):
    """Generate the S-curve plot with proper error handling"""
    if progress_df is None or progress_df.empty:
        st.warning("No valid data available for plotting")
        return None
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Plot curves
    ax.plot(progress_df['date'], progress_df['planned'], label='Planned', color='blue')
    ax.plot(progress_df['date'], progress_df['actual'], label='Actual', color='orange')
    
    # Add today line
    ax.axvline(x=today, color='red', linestyle='--', label='Today')
    
    # Calculate delay if possible
    if not progress_df[progress_df['actual'].notna()].empty:
        last_actual = progress_df[progress_df['actual'].notna()]['actual'].iloc[-1]
        
        # Interpolate planned progress at today's date
        planned_today = np.interp(
            pd.to_numeric(today),
            pd.to_numeric(progress_df['date']),
            progress_df['planned']
        )
        
        delay = last_actual - planned_today
        ax.text(today, 5, f"Delay: {delay:.1f}%", rotation=90, va='bottom', ha='right', color='red')
    
    # Format plot
    ax.set_title(f'Document Progress S-Curve (Total Weight: {total_weight:.1f})')
    ax.set_xlabel('Date')
    ax.set_ylabel('Progress (%)')
    ax.legend(loc='upper left')
    ax.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    return fig

def main():
    st.set_page_config(layout="wide")
    st.title("📊 Document Progress S-Curve Analysis")
    
    with st.expander("ℹ️ About this tool"):
        st.markdown("""
        This tool analyzes document progress by comparing planned vs actual milestones.
        - Each document goes through 4 phases (upload, review, reply, final)
        - Each phase contributes 25% to completion
        - The curve shows cumulative progress over time
        """)
    
    uploaded_file = st.file_uploader("Upload your document register (Excel file)", type=['xlsx'])
    
    if uploaded_file is not None:
        with st.spinner("Processing data..."):
            df = load_data(uploaded_file)
            
            if df is not None:
                st.success("Data loaded successfully!")
                
                # Show data preview
                if st.checkbox("Show raw data preview"):
                    st.dataframe(df.head())
                
                # Get parameters
                col1, col2 = st.columns(2)
                with col1:
                    review_days = st.number_input("Expected review duration (days)", 
                                                 min_value=1, value=10)
                with col2:
                    final_days = st.number_input("Expected finalization duration (days)", 
                                                min_value=1, value=10)
                
                # Calculate progress
                progress_df, total_weight, today = calculate_progress(df, review_days, final_days)
                
                if progress_df is not None:
                    # Display metrics
                    st.subheader("Progress Summary")
                    
                    if not progress_df[progress_df['actual'].notna()].empty:
                        last_actual = progress_df[progress_df['actual'].notna()]['actual'].iloc[-1]
                        planned_today = np.interp(
                            pd.to_numeric(today),
                            pd.to_numeric(progress_df['date']),
                            progress_df['planned']
                        )
                        delay = last_actual - planned_today
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Current Progress", f"{last_actual:.1f}%")
                        col2.metric("Planned Progress", f"{planned_today:.1f}%")
                        col3.metric("Delay", f"{delay:.1f}%", 
                                  delta_color="inverse" if delay < 0 else "normal")
                    
                    # Plot S-curve
                    st.subheader("Progress S-Curve")
                    fig = plot_s_curve(progress_df, total_weight, today)
                    if fig:
                        st.pyplot(fig)
                        
                        # Add download button
                        buf = BytesIO()
                        fig.savefig(buf, format="png", dpi=300)
                        st.download_button(
                            label="Download S-Curve",
                            data=buf.getvalue(),
                            file_name="document_progress_scurve.png",
                            mime="image/png"
                        )

if __name__ == "__main__":
    main()
