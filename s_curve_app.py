import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, timedelta

def load_data(uploaded_file):
    """Load and preprocess the Excel file"""
    df = pd.read_excel(uploaded_file)
    
    # Convert date columns to datetime
    date_cols = ['Planned Issuance date', 'Date Data Upload by EPC', 
                 'Actual Fichtner Review', 'Actual EPC reply Date', 'Final Issuance']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Fill weights with 1 if empty
    df['Weight'] = df['Weight'].fillna(1)
    
    return df

def calculate_curves(df, review_days, final_days):
    """Calculate planned and actual progress curves"""
    today = datetime.now()
    
    # Calculate planned dates
    df['Planned_Review'] = df['Planned Issuance date'] + timedelta(days=review_days)
    df['Planned_Final'] = df['Planned_Review'] + timedelta(days=final_days)
    
    # Get all unique dates from both curves
    all_dates = set()
    
    # Planned dates
    planned_dates = []
    planned_dates.extend(df['Planned Issuance date'].dropna().tolist())
    planned_dates.extend(df['Planned_Review'].dropna().tolist())
    planned_dates.extend(df['Planned_Final'].dropna().tolist())
    all_dates.update(planned_dates)
    
    # Actual dates
    actual_dates = []
    actual_dates.extend(df['Date Data Upload by EPC'].dropna().tolist())
    actual_dates.extend(df['Actual Fichtner Review'].dropna().tolist())
    actual_dates.extend(df['Actual EPC reply Date'].dropna().tolist())
    actual_dates.extend(df['Final Issuance'].dropna().tolist())
    all_dates.update(actual_dates)
    
    # Convert to sorted list
    all_dates = sorted([d for d in all_dates if pd.notnull(d)])
    
    # Calculate cumulative progress for planned curve
    planned_progress = []
    total_weight = df['Weight'].sum()
    
    for date in all_dates:
        # Count documents that have reached each milestone by this date
        uploaded = df[df['Planned Issuance date'] <= date]['Weight'].sum()
        reviewed = df[df['Planned_Review'] <= date]['Weight'].sum()
        finalized = df[df['Planned_Final'] <= date]['Weight'].sum()
        
        # Each milestone is worth 1/4 of the document's weight (upload = 0.25, review = 0.5, final = 1)
        progress = (uploaded * 0.25 + reviewed * 0.25 + finalized * 0.5) / total_weight * 100
        planned_progress.append(progress)
    
    # Calculate cumulative progress for actual curve
    actual_progress = []
    
    for date in all_dates:
        if date > today:
            # Future dates shouldn't be counted in actual progress
            actual_progress.append(np.nan)
            continue
            
        # Count documents that have reached each milestone by this date
        uploaded = df[df['Date Data Upload by EPC'] <= date]['Weight'].sum()
        reviewed = df[df['Actual Fichtner Review'] <= date]['Weight'].sum()
        replied = df[df['Actual EPC reply Date'] <= date]['Weight'].sum()
        finalized = df[df['Final Issuance'] <= date]['Weight'].sum()
        
        # Each step is worth 1/4 of the document's weight
        progress = (uploaded * 0.25 + reviewed * 0.25 + replied * 0.25 + finalized * 0.25) / total_weight * 100
        actual_progress.append(progress)
    
    # Create DataFrame with results
    progress_df = pd.DataFrame({
        'Date': all_dates,
        'Planned': planned_progress,
        'Actual': actual_progress
    })
    
    return progress_df, today

def plot_s_curve(progress_df, today):
    """Plot the S-curve with Matplotlib"""
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Plot curves
    ax.plot(progress_df['Date'], progress_df['Planned'], label='Planned', color='blue')
    ax.plot(progress_df['Date'], progress_df['Actual'], label='Actual', color='orange')
    
    # Add today line
    ax.axvline(x=today, color='red', linestyle='--', label='Today')
    
    # Calculate and display delay
    last_actual = progress_df[progress_df['Actual'].notna()]['Actual'].iloc[-1]
    planned_today = np.interp(pd.to_numeric(today), 
                             pd.to_numeric(progress_df['Date']), 
                             progress_df['Planned'])
    delay = last_actual - planned_today
    
    ax.text(today, 5, f"Delay: {delay:.1f}%", rotation=90, va='bottom', ha='right', color='red')
    
    # Formatting
    ax.set_title('Document Progress S-Curve')
    ax.set_xlabel('Date')
    ax.set_ylabel('Progress (%)')
    ax.legend()
    ax.grid(True)
    
    # Rotate x-axis labels for better readability
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    return fig

def main():
    st.title("Document Register S-Curve Analysis")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Document Register Excel File", type=['xlsx'])
    
    if uploaded_file is not None:
        # Load data
        df = load_data(uploaded_file)
        
        # Display basic stats
        st.subheader("Document Register Summary")
        st.write(f"Total documents: {len(df)}")
        st.write(f"Total weight: {df['Weight'].sum():.1f}")
        
        # User inputs
        st.subheader("S-Curve Parameters")
        col1, col2 = st.columns(2)
        with col1:
            review_days = st.number_input("Review Duration (days)", min_value=1, value=10)
        with col2:
            final_days = st.number_input("Final Issuance Duration (days)", min_value=1, value=10)
        
        # Calculate curves
        progress_df, today = calculate_curves(df, review_days, final_days)
        
        # Plot
        st.subheader("Progress S-Curve")
        fig = plot_s_curve(progress_df, today)
        st.pyplot(fig)
        
        # Display current status
        last_actual = progress_df[progress_df['Actual'].notna()]['Actual'].iloc[-1]
        planned_today = np.interp(pd.to_numeric(today), 
                                 pd.to_numeric(progress_df['Date']), 
                                 progress_df['Planned'])
        delay = last_actual - planned_today
        
        st.metric("Current Actual Progress", f"{last_actual:.1f}%")
        st.metric("Planned Progress for Today", f"{planned_today:.1f}%")
        st.metric("Delay", f"{delay:.1f}%", delta_color="inverse")

if __name__ == "__main__":
    main()
