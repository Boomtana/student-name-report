import pandas as pd
import sys
import os

def extract_data(file_path):
    try:
        if not os.path.exists(file_path):
            print(f"Error: File '{file_path}' not found.")
            return None

        # Load the Excel file without headers
        # Column B is index 1, E is index 4, H is index 7
        df = pd.read_excel(file_path, header=None)
        
        # Select columns B, E, and H
        # Rows start at index 9 (which is Excel row 10)
        extracted_df = df.iloc[9:, [1, 4, 7]]
        
        # Rename columns for clarity (optional, based on headers in row 10)
        extracted_df.columns = ['No.', 'Student ID', 'Name-Surname']
        
        # Drop rows that are completely empty in our selected columns
        extracted_df = extracted_df.dropna(how='all')
        
        # Reset index for a clean output
        extracted_df = extracted_df.reset_index(drop=True)
        
        return extracted_df

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 extract_student_data.py <filename.xls>")
        sys.exit(1)

    file_name = sys.argv[1]
    data = extract_data(file_name)
    
    if data is not None:
        print(f"--- Extracted Student Data from {file_name} ---")
        print(data.to_string(index=False))
        
        # Generate output name based on input name
        base_name = os.path.splitext(file_name)[0]
        output_csv = f"{base_name}_extracted.csv"
        
        data.to_csv(output_csv, index=False, encoding='utf-8-sig')
        print(f"\nData also saved to {output_csv}")
