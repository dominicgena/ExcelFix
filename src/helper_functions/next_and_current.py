from config.main import debug_mode
from datetime import datetime
from src.helper_functions.functions import debug_log
from sheet_to_array import load_spreadsheet_to_array
from src.helper_functions.functions import excel_time_to_string



def get_current_day_time():
    """Returns the current day as a string and time in Excel format."""
    now = datetime.now()
    current_day = now.strftime('%A')
    current_time = (now.hour + (now.minute / 60)) / 24  # Convert to Excel time format

    debug_log(f"Current day: {current_day}", debug_mode)
    debug_log(f"Current time (Excel format): {current_time}", debug_mode)

    return current_day, current_time


def get_next_lesson(lessons, current_day, current_time):
    lessons = load_spreadsheet_to_array(debug_mode)
    """Returns the next lesson based on the current day and time."""
    days_of_week = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    current_day_index = days_of_week.index(current_day)

    for i in range(7):
        day_index = (current_day_index + i) % 7
        day = days_of_week[day_index]
        current_day_lessons = lessons.get(day, [])

        for lesson in current_day_lessons:
            lesson_time = float(lesson[7])  # Start time is in the 8th column
            if day == current_day and lesson_time > current_time:
                return lesson
            elif day != current_day:
                return lesson

    return None


def get_current_lesson(lessons, current_day, current_time):
    lessons = load_spreadsheet_to_array(debug_mode)
    current_day_lessons = lessons.get(current_day, [])
    current_lesson = None
    for lesson in current_day_lessons:
        try:
            start_time = float(lesson[7])
            end_time = float(lesson[8])
        except (ValueError, TypeError) as e:
            debug_log(f"Error converting lesson times to float: {e}", debug_mode)
            continue

        debug_log(f"Checking lesson: {lesson}", debug_mode)
        debug_log(f"Start time: {start_time}, End time: {end_time}, Current time: {current_time}", debug_mode)

        if start_time <= current_time <= end_time:
            current_lesson = lesson
            break

    return current_lesson

# current_day, current_time = get_current_day_time()
#
# next_lesson = get_next_lesson(lessons, current_day, current_time)
# current_lesson = get_current_lesson(lessons, current_day, current_time)
# # reset all highlighted rows
# if next_lesson:
#     print(f"Next lesson is: {next_lesson[1]} at {excel_time_to_string(next_lesson[7])} on {next_lesson[3]}")
#     # highlight_row(next_lesson[0], "#FF0000")
# else:
#     print("No upcoming lessons.")
#
# if current_lesson:
#     print(
#         f"Current lesson is: {current_lesson[1]} at {excel_time_to_string(current_lesson[7])} to {excel_time_to_string(current_lesson[8])}")
# else:
#     print("No current lesson is in session.")