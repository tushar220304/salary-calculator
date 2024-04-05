import openpyxl
from datetime import datetime, timedelta, time

WAGE_PER_HOUR = 54
WORKING_HOUR = timedelta(hours=8, minutes=30).total_seconds() / 3600

def time_format(time_obj):
    time_obj = datetime.strptime(str(time_obj), '%H:%M:%S').time()
    # time_obj = datetime.combine(datetime.min, time_obj)
    return time_obj

def total_working_hour(end_time, start_time):
    end_time = datetime.combine(datetime.min, end_time)
    start_time = datetime.combine(datetime.min, start_time)
    duration = end_time - start_time
    duration_hours = duration.total_seconds() / 3600
    return round(duration_hours, 2)

def overtime_calculator(total_working_duration):
    if total_working_duration > WORKING_HOUR:
        overtime = total_working_duration - WORKING_HOUR
    else:
        overtime = 0.0
    return round(overtime, 2)

def wage_calculator(working_hour):
    return round(working_hour*WAGE_PER_HOUR , 2)

def open_excel(file):
    workbook = openpyxl.load_workbook(file, data_only=True)

    worksheet = workbook.active

    basic_details = {'Month': worksheet['B1'].value, 'Name': worksheet['C2'].value}

    column_index = 1
    start_row = 4


    dates = []
    total_salary = 0

    for row in worksheet.iter_rows(min_row=start_row, min_col=column_index, max_col=4):
        row_empty = all(cell.value is None for cell in row)
        if row_empty:
            continue
        single_row = []
        row_index = 1
        for cell in row:
            if isinstance(cell.value, datetime):
                single_row.append(time_format(cell.value))
            elif cell.value is None and row_index in [3, 4]:
                single_row.append(time_format(time()))
            else:
                single_row.append(cell.value)
            row_index += 1
        single_row.append(total_working_hour(single_row[3], single_row[2]))
        single_row.append(overtime_calculator(single_row[4]))
        single_row.append(wage_calculator(single_row[4]))
        total_salary += single_row[6]
        print(single_row)
        dates.append(single_row)
    return dates, round(total_salary), basic_details




