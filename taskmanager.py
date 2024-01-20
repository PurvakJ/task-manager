import time
import datetime
from plyer import notification
import pandas as pd
from xlwt import Workbook
import xlrd
from openpyxl import load_workbook
import pyautogui

file_path = 'D:/downloads/Reminder.xlsx'
backup_path = 'C:/Users/Dell/PycharmProjects/jarvisAI/pythonProject1/xlwt_backup.xls'

# Initialize a dictionary to track reminder states
reminder_states = {}

while True:
    dataframe1 = pd.read_excel(file_path)
    workbook = xlrd.open_workbook(backup_path)
    rows = workbook.sheet_by_index(0)

    # Example: Print the value of the cell in the first row and first column
    rows = rows.cell_value(0, 0)
    rows = int(rows)
    print(rows)

    for index, row in dataframe1.iterrows():
        title = row['TITLE']
        message = row['MESSAGE']
        time_interval = row['TIME']

        try:
            target_time = datetime.datetime.strptime(str(time_interval), "%H:%M:%S").time()
        except ValueError:
            print(f"Invalid time format in row {index + 2}. Please enter time in HH:MM:SS format.")
            continue

        current_datetime = datetime.datetime.now()
        target_datetime = datetime.datetime.combine(current_datetime.date(), target_time)

        # Check for reminder condition and cooldown, only notifying once per reminder
        if current_datetime >= target_datetime and title not in reminder_states:
            notification.notify(
                title="ALERT: " + title,
                message=message,
                timeout=10
            )
            reminder_states[title] = True  # Mark reminder as notified

            try:
                # Try loading the existing Excel file
                df = pd.read_excel(backup_path)
            except FileNotFoundError:
                # If the file doesn't exist, create a new DataFrame
                df = pd.DataFrame()

            # Create a backup with xlwt
            wb = Workbook()
            sheet1 = wb.add_sheet('Sheet 1')

            sheet1.write(rows, 0, title)
            sheet1.write(rows, 1, message)
            sheet1.write(rows, 2, time_interval)
            sheet1.write(rows, 3, 'HIT')
            sheet1.write(rows, 4, current_datetime.strftime("%Y-%m-%d %H:%M:%S"))
            sheet1.write(0, 0, rows + 1)
            try:
                wb.save('C:/Users/Dell/PycharmProjects/jarvisAI/pythonProject1/xlwt_backup.xls')
            except PermissionError as e:
                print(f"PermissionError: {e}")
            except Exception as e:
                print(f"An error occurred: {e}")
    time.sleep(1)
