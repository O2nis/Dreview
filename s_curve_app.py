import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import chrono
import io
import re

# Function to safely parse dates
def parse_date(date_str):
    if pd.isna(date_str) or date_str == '':
        return None
    try:
        parsed = chrono.parse_date(str(date_str))
        return parsed if parsed else None
    except:
        return None

# Function to verify column by partial name
def find_column_index(columns, patterns):
    for pattern in patterns:
        for i, col in enumerate(columns):
            if re.search(pattern, str(col), re.IGNORECASE):
                return i
    return None

# Function to process CSV and calculate S-curves
def process_data(df, start_date, review_days, final_issuance_days, today):
    # Verify CSV has enough columns
    if len(df.columns) < 23:
        raise ValueError("CSV must have at least 23 columns (up to column W).")

    # Define column indices (0-based)
    col_indices = {
        'planned': 18,  # Column S
        'actual_issuance': 19,  # Column T
        'review': 20,  # Column U
        'reply': 21,  # Column V
        'final_issuance': 22,  # Column W
        'weight': 14  # Column O
    }

    # Verify columns with partial name matching
    expected_patterns = {
        'planned': [r'planned.*issuance'],
        'actual_issuance': [r'date.*data.*upload', r'actual.*issuance'],
        'review': [r'actual.*fichtner.*review', r'review.*date'],
        'reply': [r'actual.*epc.*reply', r'reply.*date'],
        'final_issuance': [r'final.*issuance'],
        'weight': [r'weight']
    }
    for key, index in col_indices.items():
        col_name = df.columns[index]
        patterns = expected_patterns[key]
        if not any(re.search(p, str(col_name), re.IGNORECASE) for p in patterns):
            st.warning(f"Column {col_name} at index {index} may not match expected {key} column.")

    # Create a mapping of column indices to data
    date_columns = [
        ('planned', col_indices['planned']),
        ('actual_issuance', col_indices['actual_issuance']),
        ('review', col_indices['review']),
        ('reply', col_indices['reply']),
        ('final_issuance', col_indices['final_issuance'])
    ]

    # Parse dates
    for _, col_idx in date_columns:
        df.iloc[:, col_idx] = df.iloc[:, col_idx].apply(parse_date)
    
    # Handle weights
    df['Weight'] = df.iloc[:, col_indices['weight']].apply(
        lambda x: 1.0 if pd.isna(x) or x == '' else float(x)
    )
    total_weight = df['Weight'].sum() * 4  # Each document has 4 milestones
    
    # Actual S-curve: Count completed milestones up to today
    actual_data = []
    dates = []
    for _, row in df.iterrows():
        weight_per_milestone = row['Weight'] / 4
        for col_key, col_idx in date_columns[1:]:  # Skip planned
            date = row.iloc[col_idx]
            if date and date <= today:
                dates.append((date, weight_per_milestone))
    
    # Sort dates and calculate cumulative progress
    dates.sort(key=lambda x: x[0])
    actual_cumulative = []
    cum_weight = 0
    last_date = start_date
    for date, weight in dates:
        if date > today:
            break
        if date > last_date:
            actual_cumulative.append((last_date, cum_weight / total_weight * 100))
            last_date = date
        cum_weight += weight
    actual_cumulative.append((min(today, last_date), cum_weight / total_weight * 100))
    
    # Planned S-curve
    planned_data = []
    for _, row in df.iterrows():
        weight_per_milestone = row['Weight'] / 4
        planned_date = row.iloc[col_indices['planned']]
        if planned_date:
            # Issuance
            planned_data.append((planned_date, weight_per_milestone))
            # Review
            review_date = planned_date + timedelta(days=review_days)
            planned_data.append((review_date, weight_per_milestone))
            # Reply (midpoint approximation)
            reply_date = planned_date + timedelta(days=review_days + final_issuance_days / 2)
            planned_data.append((reply_date, weight_per_milestone))
            # Final issuance
            final_date = planned_date + timedelta(days=review_days + final_issuance_days)
            planned_data.append((final_date, weight_per_milestone))
    
    # Sort and calculate cumulative
    planned_data.sort(key=lambda x: x[0])
    planned_cumulative = []
    cum_weight = 0
    last_date = start_date
    for date, weight in planned_data:
        if date > last_date:
            planned_cumulative.append((last_date, cum_weight / total_weight * 100))
            last_date = date
        cum_weight += weight
    planned_cumulative.append((last_date, cum_weight / total_weight * 100))
    
    # Find delay at today
    actual_progress = actual_cumulative[-1][1] if actual_cumulative else 0
    planned_at_today = next((p for d, p in planned_cumulative if d >= today), planned_cumulative[-1][1])
    delay = planned_at_today - actual_progress
    
    return actual_cumulative, planned_cumulative, delay

# Plot S-curve
def plot_s_curves(actual_data, planned_data, today, delay, expected_end_date):
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # Plot curves
    if actual_data:
        dates, values = zip(*actual_data)
        ax.plot(dates, values, label='Actual Progress', color='blue')
    
    if planned_data:
        dates, values = zip(*planned_data)
        ax.plot(dates, values, label='Planned Progress', color='green')
    
    # Add vertical lines
    ax.axvline(x=today, color='red', linestyle='--', label=f'Today ({today.strftime("%b %d, %Y")})')
    ax.axvline(x=expected_end_date, color='purple', linestyle='--', 
               label=f'Expected End ({expected_end_date.strftime("%b %d, %Y")})')
    
    # Customize plot
    ax.set_xlabel('Date')
    ax.set_ylabel('Progress (%)')
    ax.set_title(f'S-Curve: Actual vs Planned (Delay: {delay:.2f}%)')
    ax.legend()
    ax.grid(True)
    
    # Format x-axis dates
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    plt.xticks(rotation=45)
    
    plt.tight_layout()
    
    # Save to bytes buffer
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    return buf

# Streamlit app
st.title('Document Register S-Curve')

# User inputs
start_date = st.date_input('Start Date', value=datetime(2024, 8, 1))
review_days = st.number_input('Review Days', min_value=1, value=10)
final_issuance_days = st.number_input('Final Issuance Days', min_value=1, value=10)

# Read CSV
csv_content = st.file_uploader('Upload CSV', type='csv')
if csv_content:
    try:
        df = pd.read_csv(csv_content)
        
        # Process data
        today = datetime(2025, 5, 1)
        actual_data, planned_data, delay = process_data(df, start_date, review_days, final_issuance_days, today)
        
        # Find expected end date
        expected_end_date = max([d for d, _ in planned_data], default=today)
        
        # Plot and display
        buf = plot_s_curves(actual_data, planned_data, today, delay, expected_end_date)
        st.image(buf)
    except Exception as e:
        st.error(f"Error processing CSV: {str(e)}")
else:
    st.write('Please upload a CSV file.')
