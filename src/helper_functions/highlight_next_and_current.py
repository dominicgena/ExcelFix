from highlight_cells import highlight_row
from sheet_to_array import lessons
from next_and_current import get_current_day_time, get_next_lesson, get_current_lesson
from src.helper_functions.functions import excel_time_to_string, cls
import time

def highlight():
    # reset all highlighted rows before starting
    for day_lessons in lessons.values():
        for lesson in day_lessons:
            highlight_row(lesson[0], "reset")
    r = None
    s = None
    while True:
        current_day, current_time = get_current_day_time()
        next_lesson = get_next_lesson(lessons, current_day, current_time)
        current_lesson = get_current_lesson(lessons, current_day, current_time)

        if next_lesson:
            # cls()
            print(f"Next lesson is: {next_lesson[1]} at {excel_time_to_string(next_lesson[7])} on {next_lesson[3]}")
            # print(f"Next lesson is: {next_lesson[1]} at {next_lesson[7]} on {next_lesson[3]}")
            highlight_row(next_lesson[0], "#90EE90") # highlight next lesson light green
            r = next_lesson[0]

        else:
            highlight_row(r, "reset")
            print("No upcoming lessons.")

        if current_lesson:
            print(f"Current lesson is: {current_lesson[1]} at {excel_time_to_string(current_lesson[7])} to {excel_time_to_string(current_lesson[8])}")
            # print(f"Current lesson is: {current_lesson[1]} at {current_lesson[7]} to {current_lesson[8]}")
            highlight_row(current_lesson[0], "#EEE880")
            s = current_lesson[0]
        else:
            highlight_row(s, "reset") # huh?
            print("No current lesson is in session.")

        time.sleep(5)

highlight()
# region reset
# for day_lessons in lessons.values():
#     for lesson in day_lessons:
#         highlight_row(lesson[0], "reset")
# endregion reset