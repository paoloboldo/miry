import numpy as np
import openpyxl
import pandas as pd
from datetime import datetime, timedelta

def complete_joint_cells(array):
    for i, value in enumerate(array):
        if str(value) == "nan":
            array[i] = array[i-1]
    return array


def first_last_indices(array):
    first_i = None
    last_i = None
    for i, value in enumerate(array):
        if str(value) != "nan":
            last_i = i
            if first_i is None:
                first_i = i
            continue
        else:
            if first_i is not None:
                break
    return [first_i, last_i]

def get_week_sheet_names(file):
    wb = openpyxl.load_workbook(file)
    sheet_names = wb.sheetnames
    weeks = []
    for name in sheet_names:
        if "Sett" in name:
            weeks.append(name)
    # Sort numerically by extracting the number after "Sett"
    weeks.sort(key=lambda x: int(x.replace("Sett", "")))
    return weeks

def calculate_time_interval(times):
    """Calculate the interval between consecutive time entries in minutes."""
    if len(times) < 2:
        return 30  # Default to 30 minutes if can't calculate
    
    try:
        # Convert first two times to datetime objects and find the difference
        time1 = datetime.strptime(str(times[0]), "%H:%M")
        time2 = datetime.strptime(str(times[1]), "%H:%M")
        diff_minutes = (time2 - time1).total_seconds() / 60
        return diff_minutes
    except:
        return 30  # Default to 30 minutes if conversion fails

def add_minutes_to_time(time_str, minutes):
    """Add specified minutes to a time string."""
    try:
        time_obj = datetime.strptime(str(time_str), "%H:%M")
        new_time = time_obj + timedelta(minutes=minutes)
        return new_time.strftime("%H:%M")
    except:
        return f"{time_str}+{minutes}min"  # Fallback format if conversion fails

def main(file):
    # Get the names of the sheets in the workbook
    weeks = get_week_sheet_names(file)

    if not weeks:
        raise ValueError("No week sheets found in the workbook.")
    
    sheet_name = weeks[0]
    # print("sheet_name", sheet_name)
    df = pd.read_excel(file, sheet_name=sheet_name)
    values = df.values

    worker = values[1]  # nome dipendente
    first_w, last_w = first_last_indices(worker)
    worker = worker[first_w: last_w + 1]
    # print("worker", worker)

    month_day = complete_joint_cells(df.columns.to_numpy()[first_w: last_w + 1])  # giorno del mese
    # print("month_day", month_day)

    week_day = complete_joint_cells(values[0, first_w: last_w + 1])  # giorno della settimana
    # print("week_day", week_day)

    times = values[:,0]
    first_t, last_t = first_last_indices(times)
    times = times[first_t: last_t + 1]
    # print("times", times)
    
    # Calculate the time interval in minutes
    interval_minutes = calculate_time_interval(times)
    print(f"Detected time interval: {interval_minutes} minutes")

    values = values[first_t: last_t + 1, first_w: last_w + 1]
    # print("values", values)
    
    # Identify unique employees and number of days
    unique_employees = list(set(worker))
    num_employees = len(unique_employees)
    num_days = len(worker) // num_employees
    
    # Process each day
    for day_idx in range(num_days):
        day_start_col = day_idx * num_employees
        day_end_col = day_start_col + num_employees - 1
        day_name = f"{week_day[day_start_col]} {month_day[day_start_col]}"
        
        print(f"\nSchedules for {day_name}:")
        
        # Process each employee for this day
        for col_idx in range(day_start_col, day_end_col + 1):
            emp = worker[col_idx]
            
            # Extract working hours for this employee on this day
            shifts = []
            start_time = None
            other_hours = {}  # Dictionary to track hours for other types of entries
            
            for time_idx, time_slot in enumerate(times):
                value = values[time_idx, col_idx]
                value_str = str(value).strip() if not pd.isna(value) else "nan"
                
                # Skip NaN values and end any ongoing "L" shift
                if value_str == "nan":
                    if start_time is not None:
                        # End the current shift
                        end_time = add_minutes_to_time(times[time_idx-1], interval_minutes)
                        shifts.append(f"{start_time}-{end_time}")
                        start_time = None
                    continue
                    
                # For "L" values, track shift intervals
                if value_str.upper() == "L":
                    if start_time is None:
                        # Start a new shift
                        start_time = time_slot
                else:
                    # End any ongoing "L" shift
                    if start_time is not None:
                        end_time = add_minutes_to_time(times[time_idx-1], interval_minutes)
                        shifts.append(f"{start_time}-{end_time}")
                        start_time = None
                        
                    # Count hours for non-"L" values
                    if value_str not in other_hours:
                        other_hours[value_str] = 0
                    other_hours[value_str] += interval_minutes / 60
            
            # Handle case where shift extends to the end of the data
            if start_time is not None:
                end_time = add_minutes_to_time(times[-1], interval_minutes)
                shifts.append(f"{start_time}-{end_time}")
            
            # Build the output string
            output_parts = []
            if shifts:
                shift_lines = ["L:"] + shifts
                output_parts.append("\n    ".join(shift_lines))
            
            for value_type, hours in other_hours.items():
                output_parts.append(f"{value_type}: {hours:.2f} ore")
            
            schedule_str = "\n  ".join(output_parts) if output_parts else "No entries"
            print(f"  {emp}:\n  {schedule_str}")


if __name__ == "__main__":
    file = r"TURNI GENNAIO 2025.xlsx"
    main(file)