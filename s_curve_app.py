import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, timedelta
import openpyxl

def load_data(uploaded_file):
    """Load and preprocess the Excel file"""
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    # Clean column names by removing extra spaces and making consistent
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
        if not found:
            st.warning(f"Could not find column matching: {standard_name}")
    
    # Convert date columns to datetime, handling Excel serial dates
    for col in date_cols:
        try:
            # First attempt to convert assuming standard date formats
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
            
            # If there are still NaTs and the column has numeric values, try Excel serial dates
            if df[col].isna().any() and pd.api.types.is_numeric_dtype(df[col]):
                mask = df[col].notna() & df[col].apply(lambda x: isinstance(x, (int, float)))
                df.loc[mask, col] = pd.to_datetime(df.loc[mask, col], unit='D', origin='1899-12-30', errors='coerce')
            
            # Debug: Show unique values if conversion issues persist
            if df[col].isna().all() and not df[col].empty:
                st.warning(f"Failed to convert column '{col}' to datetime. Unique values: {df[col].unique()}")
        except Exception as e:
            st.error(f"Error converting column '{col}' to datetime: {str(e)}")
            st.write(f"Unique values in '{col}': {df[col].unique()}")
            raise
    
    # Fill weights with 1 if empty
    if 'Weight' in df.columns:
        df['Weight'] = df['Weight'].fillna(1)
    else:
        df['Weight'] = 1  # Default weight if column doesn't exist
    
    return df

def calculate_curves(df, review_days, final_days):
    """Calculate planned and actual progress curves"""
    today = datetime(2025, 5, 1)  # Fixed as per instruction
    
    # Initialize variables
    total_weight = df['Weight'].sum()
    all_dates = set()
    
    # Calculate planned dates
    if 'Planned Issuance date' in df.columns and df['Planned Issuance date'].notna().any():
        df['Planned_Review'] = df['Planned Issuance date'] + timedelta(days=review_days)
        df['Planned_Final'] = df['Planned_Review'] + timedelta(days=final_days)
        
        # Collect planned dates
        planned_dates = []
        for col in ['Planned Issuance date', 'Planned_Review', 'Planned_Final']:
            valid_dates = df[col].dropna().tolist()
            if not valid_dates:
                st.warning(f"No valid dates found in '{col}' for planned S-curve")
            planned_dates.extend(valid_dates)
        all_dates.update(planned_dates)
    else:
        st.error("Missing or empty 'Planned Issuance date' column. Cannot calculate planned S-curve.")
        return pd.DataFrame(), today
    
    # Collect actual dates
    actual_dates = []
    actual_cols = ['Date Data Upload by EPC', 'Actual Fichtner Review', 
                   'Actual EPC reply Date', 'Final Issuance']
    for col in actual_cols:
        if col in df.columns and df[col].notna().any():
            valid_dates = df[df[col] <= today][col].dropna().tolist()
            if not valid_dates:
                st.warning(f"No valid dates found in '{col}' up to today for actual S-curve")
            actual_dates.extend(valid_dates)
        else:
            st.warning(f"Column '{col}' missing or empty for actual S-curve")
    all_dates.update(actual_dates)
    
    # Handle case where no valid dates are found
    if not all_dates:
        st.error("No valid dates found for planned or actual S-curves. Please check your data.")
        return pd.DataFrame(), today
    
    # Convert to sorted list
    all_dates = sorted(all_dates)
    
    # Calculate cumulative progress for planned curve
    planned_progress = []
    for date in all_dates:
        uploaded = df[df['Planned Issuance date'] <= date]['Weight'].sum() if 'Planned Issuance date' in df.columns else 0
        reviewed = df[df['Planned_Review'] <= date]['Weight'].sum() if 'Planned_Review' in df.columns else 0
        finalized = df[df['Planned_Final'] <= date]['Weight'].sum() if 'Planned_Final' in df.columns else 0
        progress = (uploaded * 0.25 + reviewed * 0.25 + finalized * 0.5) / total_weight * 100
        planned_progress.append(progress)
    
    # Calculate cumulative progress for actual curve
    actual_progress = []
    for date in all_dates:
        if date > today:
            actual_progress.append(np.nan)
            continue
        uploaded = df[df['Date Data Upload by EPC'] <= date]['Weight'].sum() if 'Date Data Upload by EPC' in df.columns and df['Date Data Upload by EPC'].notna().any() else 0
        reviewed = df[df['Actual Fichtner Review'] <= date]['Weight'].sum() if 'Actual Fichtner Review' in df.columns and df['Actual Fichtner Review'].notna().any() else 0
        replied = df[df['Actual EPC reply Date'] <= date]['Weight'].sum() if 'Actual EPC reply Date' in df.columns and df['Actual EPC reply Date'].notna().any() else 0
        finalized = df[df['Final Issuance'] <= date]['Weight'].sum() if 'Final Issuance' in df.columns and df['Final Issuance'].notna().any() else 0
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
    
    # Calculate and display delay if we have data
    if not progress_df.empty:
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
        try:
            # Load data
            df = load_data(uploaded_file)
            
            # Display basic stats
            st.subheader("Document Register Summary")
            st.write(f"Total documents: {len(df)}")
            st.write(f"Total weight: {df['Weight'].sum():.1f}")
            
            # Show column names for debugging
            if st.checkbox("Show column names"):
                st.write("Columns in uploaded file:", df.columns.tolist())
            
            # User inputs
            st.subheader("S-Curve Parameters")
            col1, col2 = st.columns(2)
            with col1:
                review_days = st.number_input("Review Duration (days)", min_value=1, value=10)
            with col2:
                final_days = st.number_input("Final Issuance Duration (days)", min_value=1, value=10)
            
            # Calculate curves
            progress_df, today = calculate_curves(df, review_days, final_days)
            
            if not progress_df.empty:
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
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.write("Please check that your file has the required columns and valid date formats.")

if __name__ == "__main__":
    main()
