import pandas as pd
import numpy as np
import random
from datetime import timedelta

def generate_dummy_data(df, num_rows=100000):
    """Generates a dummy dataset based on the structure of an input DataFrame."""

    dummy_data = {}

    for col in df.columns:
        dummy_data[col] = []

        # Check if the column is completely empty
        if df[col].isnull().all():
            dummy_data[col] = [''] * num_rows  # Fill completely empty columns with empty strings
            continue  # Move to the next column

        if df[col].dtype == 'object':  # String/Categorical column
            unique_vals = df[col].dropna().unique()
            if len(unique_vals) > 0:
                dummy_data[col] = random.choices(unique_vals, k=num_rows)
            else:
                dummy_data[col] = [''] * num_rows # Handle categorical columns with no values
        elif df[col].dtype in ['int64', 'float64']:  # Numerical column
            min_val = df[col].min()
            max_val = df[col].max()

            if pd.isna(min_val) or pd.isna(max_val): # If all values are NaN
                dummy_data[col] = [np.nan] * num_rows # Fill column with NaNs
            else:
                if df[col].dtype == 'int64':
                    dummy_data[col] = np.random.randint(min_val, max_val + 1, size=num_rows)
                else:
                    dummy_data[col] = np.random.uniform(min_val, max_val, size=num_rows)

        elif df[col].dtype == 'datetime64[ns]':  # Datetime column
            start_date = pd.to_datetime('2021-01-01')
            end_date = pd.to_datetime('2022-12-31')
            time_between_dates = end_date - start_date
            days_between_dates = time_between_dates.days
            random_number_of_days = np.random.randint(0, days_between_dates, size=num_rows)
            dummy_data[col] = start_date + pd.to_timedelta(random_number_of_days, unit='D')
        else:
            dummy_data[col] = [''] * num_rows  # Default: fill with empty strings

    return pd.DataFrame(dummy_data)

# Load the Excel file
try:
    excel_file = pd.ExcelFile("sample_data_eda.xlsx")
    df = excel_file.parse("Sheet1")
except FileNotFoundError:
    print("Error: The file 'sample_data_eda.xlsx' was not found.  Make sure it's in the correct directory.")
    exit()
except Exception as e:
    print(f"An error occurred while reading the Excel file: {e}")
    exit()

# Generate the dummy data
try:
    dummy_df = generate_dummy_data(df)
except Exception as e:
    print(f"An error occurred during data generation: {e}")
    exit()

# Save to an XLSX file
try:
    dummy_df.to_excel("dummy_data.xlsx", index=False, engine='openpyxl')
    print("Dummy data successfully generated and saved to 'dummy_data.xlsx'")
except Exception as e:
    print(f"An error occurred while saving to XLSX: {e}")
    exit()
