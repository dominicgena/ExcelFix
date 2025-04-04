from collections import defaultdict
import xlwings as xw
from config.main import debug_mode
from src.helper_functions.functions import debug_log
from tabulate import tabulate  # Install with: pip install tabulate


def load_spreadsheet_to_array(debug_mode):
    try:
        """Loads the lesson schedule into a dictionary grouped by day, shifting columns right to include Excel row numbers."""
        wb = xw.books.active
        sheet = wb.sheets["Lesson Schedule"]
        data = sheet.used_range.value

        # Get row numbers
        row_numbers = [cell.row for cell in sheet.used_range.rows]

        # Extract headers before inserting row numbers
        headers = ["Excel Row"] + data[0]  # Add "Excel Row" to headers
        data_rows = data[1:]  # Remove headers from data

        # Remove empty rows and insert row numbers
        filtered_data = [
            [row_numbers[i + 1]] + row  # Shift row number into first column
            for i, row in enumerate(data_rows)
            if row[0] is not None and row[0] != ""
        ]

        lessons_by_day = defaultdict(list)

        # Get the column index for "Day of Week"
        day_column_index = headers.index("Day of Week")

        # Sort lessons into lists by day
        for row in filtered_data:
            day = row[day_column_index]
            lessons_by_day[day].append(row)

        # Debug mode: Print formatted table
        if debug_mode:
            debug_log("\nLesson Schedule Table:\n", debug_mode)
            debug_log(tabulate(filtered_data, headers=headers, tablefmt="grid"), debug_mode)

        return lessons_by_day
    except Exception as e:
        debug_log(f"Error loading spreadsheet to array: {e}", debug_mode)
        return {}

# Load lessons once and store the result
#
# # Assign lessons to variables without calling the function again
# sunday_lessons = lessons.get("Sunday", [])
# monday_lessons = lessons.get("Monday", [])
# tuesday_lessons = lessons.get("Tuesday", [])
# wednesday_lessons = lessons.get("Wednesday", [])
# thursday_lessons = lessons.get("Thursday", [])
# friday_lessons = lessons.get("Friday", [])
# saturday_lessons = lessons.get("Saturday", [])
