import os
import pandas as pd
from datetime import date
import tracebackdef process_financial_data(file_path):
    """Process a single Excel file"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
      
        # Preserve all original columns
        original_columns = df.columns.tolist()
      
        # Precise numerical conversions
        df['payin_amount'] = df['payin_amount'].astype(float)
        df['payout_amount'] = df['payout_amount'].astype(float)
      
        # Date column conversions
        date_columns = ['start_date', 'expiry_date', 'tp_expiry_date']
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
      
        # Replace NaN with an empty string in the specified columns
        columns_to_clean = ['c_no', 'inst_no', 'pk', 'policy_no', 'start_date', 'expiry_date', 'tp_expiry_date']
        df[columns_to_clean] = df[columns_to_clean].fillna("")        # Primary key generation
        df['pk'] = df['c_no'].astype(str) + df['inst_no'].astype(str)
      
        # Month period creation
        df['month_period'] = df['start_date'].dt.strftime('%b-%y')
      
        # Add (~) prefix to specific columns
        prefix_columns = ['c_no', 'inst_no', 'pk', 'policy_no', 'start_date', 'expiry_date', 'tp_expiry_date']
        for col in prefix_columns:
            # Convert to string first
            df[col] = df[col].astype(str)
            # Replace empty strings and 'NaT' with empty string, add prefix to other values
            df[col] = df[col].apply(lambda x: '' if x == '' or x == 'NaT' else '~' + x)
            # Additional check for ~NaT
            df[col] = df[col].replace('~NaT', '')        # Retention calculation
        df['retention'] = df['payin_amount'] - df['payout_amount']
      
        # Additional metadata columns
        df['nop'] = df['nop'].apply(lambda x: 0 if pd.notna(x) and x != 0 else x)
        df['entry_date_db'] = date.today().strftime("%Y-%m-%d")
      
        return df
    except Exception as e:
        print(f"Error processing {file_path}:")
        print(traceback.format_exc())
        return Nonedef main():
    # Create output directory
    output_dir = 'financial_processing_output'
    os.makedirs(output_dir, exist_ok=True)
  
    # Get number of files to process
    while True:
        try:
            num_files = int(input("Enter number of Excel files to process: "))
            if num_files > 0:
                break
            print("Please enter a positive number.")
        except ValueError:
            print("Invalid input. Enter a number.")
  
    # Collect file paths
    file_paths = []
    print("\nEnter FULL file paths (copy-paste from File Explorer):")
    for i in range(num_files):
        while True:
            file_path = input(f"File {i+1} path: ").strip().replace('"', '')
            if os.path.exists(file_path) and file_path.lower().endswith(('.xlsx', '.xls')):
                file_paths.append(file_path)
                break
            print("Invalid path. Ensure file exists and is an Excel file.")
  
    # Process files
    successful_files = 0
    failed_files = 0
  
    print("\n--- PROCESSING STARTED ---")
    for idx, file_path in enumerate(file_paths, 1):
        try:
            print(f"\nProcessing File {idx}: {os.path.basename(file_path)}")
          
            # Process the file
            processed_df = process_financial_data(file_path)
          
            if processed_df is not None:
                # Print row numbers every 4000 rows
                for row_idx, _ in processed_df.iterrows():
                    if (row_idx + 1) % 4000 == 0:
                        print(f"Processing Row Number: {row_idx + 1}")
              
                # Generate output filename
                filename = os.path.splitext(os.path.basename(file_path))[0]
                csv_path = os.path.join(output_dir, f'cleaned_{filename}.csv')
              
                # Save CSV
                processed_df.to_csv(csv_path, index=False)
              
                # Print summary
                print(f"\n--- Summary for {filename} ---")
                print(f"Total Records: {len(processed_df)}")
                print(f"Total Payin Amount: {processed_df['payin_amount'].sum():,.2f}")
                print(f"Total Payout Amount: {processed_df['payout_amount'].sum():,.2f}")
                print(f"Total Gross Premium: {processed_df['gross_premium'].sum():,.2f}")
                print(f"Total Retention: {processed_df['retention'].sum():,.2f}")
                print(f"CSV saved to: {csv_path}")
              
                successful_files += 1
            else:
                failed_files += 1
      
        except Exception as e:
            print(f"Unexpected error processing {file_path}:")
            print(traceback.format_exc())
            failed_files += 1
  
    # Final summary
    print("\n--- PROCESSING COMPLETE ---")
    print(f"Total Files Processed: {num_files}")
    print(f"Successful Files: {successful_files}")
    print(f"Failed Files: {failed_files}")
  
    input("\nPress Enter to exit...")if _name_ == "_main_":
    main()