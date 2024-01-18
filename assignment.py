import csv
from datetime import datetime, timedelta

def analyze_employee_data(file_path):
    with open(file_path, 'r') as file:
        reader = csv.DictReader(file)
        data = list(reader)

    # Filter out rows with empty or invalid 'Time' values
    data = [row for row in data if row['Time'] and row['Time Out']]
    
    # Sort the data by employee name and time
    try:
        data.sort(key=lambda x: (x['Employee Name'], datetime.strptime(x['Time'], '%m/%d/%Y %I:%M %p')))
    except ValueError as e:
        print(f"Error sorting data: {e}")
        return

    for i in range(len(data)-2):
        # Extract relevant information
        employee_name = data[i]['Employee Name']
        position_id = data[i]['Position ID']

        try:
            time_in = datetime.strptime(data[i]['Time'], '%m/%d/%Y %I:%M %p')
            time_out = datetime.strptime(data[i]['Time Out'], '%m/%d/%Y %I:%M %p')

            # Parse hours and minutes separately
            hours, minutes = map(int, data[i]['Timecard Hours (as Time)'].split(':'))
            hours_worked = timedelta(hours=hours, minutes=minutes)
        except ValueError as e:
            print(f"Error processing data for {employee_name}: {e}")
            continue

        # Check conditions and print information
        if (time_out - time_in).days == 1 and data[i]['Employee Name'] == data[i+1]['Employee Name']:
            print("\n" f"{employee_name} worked for 7 consecutive days at position {position_id}.")

        if 1 < (time_in - time_out).total_seconds() / 3600 < 10 and data[i]['Employee Name'] == data[i+1]['Employee Name']:
            print(f"{employee_name} has less than 10 hours between shifts at position {position_id}.")

        if hours_worked > timedelta(hours=14):
            print(f"{employee_name} worked for more than 14 hours in a single shift at position {position_id}.")

# Example usage
file_path = 'D:/visual studio/python/Assignment_Timecard.xlsx - Sheet1.csv'  # Replace with the actual file path
analyze_employee_data(file_path)
