import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import numpy as np
import openpyxl  # Added import for openpyxl

# Function to parse Excel file
@st.cache_data
def load_data(file):
    df = pd.read_excel(file, engine='openpyxl')
    date_columns = ['Planned Issuance date', 'Date Data Upload by EPC', 
                    'Actual Fichtner Review', 'Actual EPC reply Date', 'Final Issuance']
    for col in date_columns:
        df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
    return df

# Function to calculate S-curves
def calculate_s_curves(df, start_date, review_days, final_issuance_days, today):
    # Initialize weights
    df['Weight'] = df['Weight'].fillna(1.0)
    total_weight = df['Weight'].sum()
    
    # Planned S-curve: Planned issuance + review + final issuance
    planned_dates = []
    planned_weights = []
    for _, row in df.iterrows():
        if pd.notna(row['Planned Issuance date']):
            issuance_date = row['Planned Issuance date']
            review_date = issuance_date + timedelta(days=review_days)
            final_date = review_date + timedelta(days=final_issuance_days)
            weight_per_step = row['Weight'] / 4  # Each step is 1/4
            # Add steps: issuance (25%), review (25%), reply (25%), final (25%)
            planned_dates.extend([issuance_date, issuance_date, review_date, final_date])
            planned_weights.extend([weight_per_step, weight_per_step, weight_per_step, weight_per_step])
    
    # Create planned DataFrame
    planned_df = pd.DataFrame({'Date': planned_dates, 'Weight': planned_weights})
    planned_df = planned_df[planned_df['Date'].notna()]
    planned_df = planned_df.sort_values('Date')
    
    # Aggregate planned weights
    date_range = pd.date_range(start=start_date, end=planned_df['Date'].max(), freq='D')
    planned_cumulative = []
    cumulative_weight = 0
    for date in date_range:
        daily_weight = planned_df[planned_df['Date'].date == date.date()]['Weight'].sum()
        cumulative_weight += daily_weight
        planned_cumulative.append(cumulative_weight)
    
    planned_cumulative = np.array(planned_cumulative) / total_weight * 100  # Convert to percentage
    
    # Actual S-curve: Based on T, U, V, W columns
    actual_dates = []
    actual_weights = []
    for _, row in df.iterrows():
        weight_per_step = row['Weight'] / 4
        for col in ['Date Data Upload by EPC', 'Actual Fichtner Review', 
                    'Actual EPC reply Date', 'Final Issuance']:
            if pd.notna(row[col]) and row[col].date() <= today.date():
                actual_dates.append(row[col])
                actual_weights.append(weight_per_step)
    
    # Create actual DataFrame
    actual_df = pd.DataFrame({'Date': actual_dates, 'Weight': actual_weights})
    actual_df = actual_df[actual_df['Date'].notna()]
    actual_df = actual_df.sort_values('Date')
    
    # Aggregate actual weights
    actual_cumulative = []
    cumulative_weight = 0
    for date in date_range:
        if date.date() > today.date():
            break
        daily_weight = actual_df[actual_df['Date'].date == date.date()]['Weight'].sum()
        cumulative_weight += daily_weight
        actual_cumulative.append(cumulative_weight)
    
    actual_cumulative = np.array(actual_cumulative) / total_weight * 100  # Convert to percentage
    
    # Calculate delay at today's date
    today_idx = (today - start_date).days
    if today_idx < len(planned_cumulative) and today_idx < len(actual_cumulative):
        delay = planned_cumulative[today_idx] - actual_cumulative[today_idx]
    else:
        delay = 0
    
    return date_range, planned_cumulative, actual_cumulative, delay, planned_df['Date'].max()

# Streamlit UI
st.title("S-Curve Analysis for Document Issuance")

# File upload
uploaded_file = st.file_uploader("Upload Excel Document Register", type=["xlsx"])
if uploaded_file is None:
    st.stop()

# Load data
df = load_data(uploaded_file)

# User inputs
start_date = st.date_input("Start Date", value=datetime(2025, 1, 1))
review_days = st.number_input("Review Days", min_value=1, value=10)
final_issuance_days = st.number_input("Final Issuance Days", min_value=1, value=10)
today = datetime(2025, 5, 1).date()  # Fixed as per instruction

# Calculate S-curves
date_range, planned_cumulative, actual_cumulative, delay, expected_end_date = calculate_s_curves(
    df, start_date, review_days, final_issuance_days, today
)

# Plot S-curves
fig, ax = plt.subplots(figsize=(10, 6))
ax.plot(date_range[:len(planned_cumulative)], planned_cumulative, label='Planned S-Curve', color='blue')
ax.plot(date_range[:len(actual_cumulative)], actual_cumulative, label='Actual S-Curve', color='red')

# Add vertical line for today
ax.axvline(x=today, color='green', linestyle='--', label="Today's Date")
ax.text(today, max(planned_cumulative) * 0.9, f'Delay: {delay:.2f}%', color='black')

# Add vertical line for expected end date
ax.axvline(x=expected_end_date, color='purple', linestyle='--', label='Expected End Date')

# Customize plot
ax.set_xlabel('Date')
ax.set_ylabel('Cumulative Completion (%)')
ax.set_title('S-Curve: Planned vs Actual Progress')
ax.legend()
ax.grid(True)

# Format x-axis dates
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
ax.xaxis.set_major_locator(mdates.AutoDateLocator())
plt.xticks(rotation=45)

# Save and display plot
plt.tight_layout()
plt.savefig('s_curve.png')
st.image('s_curve.png')