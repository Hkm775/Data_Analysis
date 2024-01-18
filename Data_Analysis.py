import pandas as pd

def analyze_excel_file(file_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Convert the 'Time' and 'Time Out' columns to datetime format
    df['Time'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])

    # Sort the DataFrame by employee and date
    df.sort_values(by=['Employee Name', 'Time'], inplace=True)

    # Reset the index after sorting
    df.reset_index(drop=True, inplace=True)

    # Step 2a: Find employees who have worked for 7 consecutive days
    consecutive_days_threshold = 7
    consecutive_days_employees = []
    for employee in df['Employee Name'].unique():
        employee_data = df[df['Employee Name'] == employee]
        consec_days = employee_data['Time'].diff().dt.days
        if any(consec_days >= consecutive_days_threshold - 1):
            consecutive_days_employees.append(employee)

    # Print the result for 2a
    print("Employees who have worked for 7 consecutive days:")
    for emp in consecutive_days_employees:
        print(f"Employee: {emp}, Position: {df[df['Employee Name'] == emp].iloc[0]['Position ID']}")

    # Step 2b: Find employees with less than 10 hours between shifts but greater than 1 hour
    min_hours_between_shifts = 1
    max_hours_between_shifts = 10
    short_break_employees = df[(df['Time'].diff().dt.total_seconds() / 3600 < max_hours_between_shifts) &
                               (df['Time'].diff().dt.total_seconds() / 3600 > min_hours_between_shifts)]['Employee Name'].unique()

    # Print the result for 2b
    print("\nEmployees with less than 10 hours between shifts but greater than 1 hour:")
    for emp in short_break_employees:
        print(f"Employee: {emp}, Position: {df[df['Employee Name'] == emp].iloc[0]['Position ID']}")

    # Step 2c: Find employees who have worked for more than 14 hours in a single shift
    max_hours_in_single_shift = 14
    long_shift_employees = df[(df['Time Out'] - df['Time']).dt.total_seconds() / 3600 > max_hours_in_single_shift]['Employee Name'].unique()

    # Print the result for 2c
    print("\nEmployees who have worked for more than 14 hours in a single shift:")
    for emp in long_shift_employees:
        print(f"Employee: {emp}, Position: {df[df['Employee Name'] == emp].iloc[0]['Position ID']}")

# Example usage:
file_path = 'Assignment_Timecard.xlsx'
analyze_excel_file(file_path)





