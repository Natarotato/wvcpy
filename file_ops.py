import pandas as pd
import os
from datetime import datetime
from .config import CONFIG

def save_data(data, sheet_name):
    """Save inverter data to an Excel file, creating directories as needed."""
    date_str = datetime.now().strftime("%Y-%m")
    file_path = os.path.join(CONFIG['data_path'], date_str, f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    
    # If file exists, load existing data to append; otherwise, create new
    if os.path.exists(file_path):
        try:
            with pd.ExcelFile(file_path) as xls:
                if sheet_name in xls.sheet_names:
                    existing_df = pd.read_excel(file_path, sheet_name=sheet_name)
                    data = pd.concat([existing_df, data], ignore_index=True)
        except Exception as e:
            print(f"Error reading existing file: {e}")
    
    # Save data to Excel
    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"✅ Data saved to '{file_path}' in sheet '{sheet_name}'")

def load_historical_data(sheet_name):
    """Load historical data for a given inverter from Excel."""
    date_str = datetime.now().strftime("%Y-%m")
    file_path = os.path.join(CONFIG['data_path'], date_str, f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")
    
    if os.path.exists(file_path):
        try:
            with pd.ExcelFile(file_path) as xls:
                if sheet_name in xls.sheet_names:
                    return pd.read_excel(file_path, sheet_name=sheet_name)
        except Exception as e:
            print(f"Error loading historical data: {e}")
    return pd.DataFrame()  # Return empty DataFrame if no data exists or error occurs