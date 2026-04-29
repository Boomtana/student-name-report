import pandas as pd
import sys
import os
from datetime import datetime, timedelta

def get_class_dates(start_date, total_periods, days_of_week):
    """
    days_of_week: list of integers (0=Mon, 1=Tue, ..., 6=Sun)
    """
    dates = []
    current_date = start_date
    while len(dates) < total_periods:
        if current_date.weekday() in days_of_week:
            dates.append(current_date.strftime('%d/%m/%y'))
        current_date += timedelta(days=1)
    return dates

def create_report(extracted_file):
    if not os.path.exists(extracted_file):
        print(f"Error: File '{extracted_file}' not found.")
        return

    df = pd.read_csv(extracted_file)
    
    # User inputs
    try:
        total_periods = int(input("Enter number of class periods: ").strip())
        days_input = input("Enter days of the week (0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri), separated by comma: ").strip()
        days_of_week = [int(d.strip()) for d in days_input.split(',')]
        start_date_str = input("Enter start date (d/m/y): ").strip()
        start_date = datetime.strptime(start_date_str, '%d/%m/%Y')
    except Exception as e:
        print(f"Invalid input ({e}). Please try again.")
        return

    dates = get_class_dates(start_date, total_periods, days_of_week)
    
    # Create the report structure
    # Header row 1: Period numbers
    # Header row 2: Dates
    
    # Initialize attendance grid with student data
    report_data = df.copy()
    
    # Add empty columns for periods
    for i in range(1, total_periods + 1):
        report_data[f'P{i}'] = ''
        
    # We'll create a multi-index or just keep it simple with rows for date and header
    # Simple approach: Create a new dataframe with two rows for headers
    
    # Constructing the full dataframe with headers
    new_rows = []
    
    # Row 1: Periods
    period_header = ['No.', 'Student ID', 'Name-Surname'] + [str(i) for i in range(1, total_periods + 1)]
    # Row 2: Dates
    date_header = ['', '', 'Date'] + dates
    
    # Data rows
    data_rows = df.values.tolist()
    
    full_data = [period_header, date_header] + data_rows
    
    # Convert to DataFrame
    final_df = pd.DataFrame(full_data)
    
    # Save to Excel
    base_name = os.path.splitext(extracted_file)[0].replace('_extracted', '')
    output_file = f"{base_name}_reportfile.xlsx"
    
    final_df.to_excel(output_file, index=False, header=False)
    print(f"Report saved to {output_file}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 create_attendance_report.py <room_extracted.csv>")
        sys.exit(1)
        
    create_report(sys.argv[1])
