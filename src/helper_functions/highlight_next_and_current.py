from config.main import debug_mode, reinit_delay
from functions import wait_for_empty_status_file
from highlight_cells import highlight_row
from sheet_to_array import load_spreadsheet_to_array
from next_and_current import get_current_day_time, get_next_lesson, get_current_lesson
from src.helper_functions.functions import excel_time_to_string, get_valid_start_seconds
import time
from datetime import datetime, timedelta

def sync_with_valid_start_seconds(valid_start_seconds):
    current_time = time.time()
    seconds = current_time % 60
    next_valid_start = min(valid_start_seconds, key=lambda x: (x - seconds) % 60)
    wait_time = (next_valid_start - seconds) % 60
    if wait_time > 0.01:
        current_time_str = datetime.now().strftime('%H:%M:%S')
        goal_start_time = (datetime.now() + timedelta(seconds=wait_time)).strftime('%H:%M:%S')
        print(f"{current_time_str} -- Waiting {wait_time:.2f} seconds for the nearest valid start time: {next_valid_start:.2f} at {goal_start_time}")
        time.sleep(wait_time)

def highlight():
    valid_start_seconds = get_valid_start_seconds(reinit_delay)
    failover_interval = 5 * 60  # Failover every 5 minutes

    # Initial synchronization
    sync_with_valid_start_seconds(valid_start_seconds)

    # reset all highlighted rows before starting
    lessons = load_spreadsheet_to_array(debug_mode)
    for day_lessons in lessons.values():
        for lesson in day_lessons:
            highlight_row(lesson[0], "reset")
    r = None
    s = None
    start_time = time.time()

    while True:
        # Failover mechanism to reset synchronization every few minutes
        if time.time() - start_time >= failover_interval:
            sync_with_valid_start_seconds(valid_start_seconds)
            start_time = time.time()

        lessons = load_spreadsheet_to_array(debug_mode)  # reload lessons
        current_day, current_time = get_current_day_time()
        next_lesson = get_next_lesson(lessons, current_day, current_time)
        current_lesson = get_current_lesson(lessons, current_day, current_time)

        if next_lesson:
            print(f"Next lesson is: {next_lesson[1]} at {excel_time_to_string(next_lesson[7])} on {next_lesson[3]}")
            if r is not None and r != next_lesson[0]:
                wait_for_empty_status_file()
                highlight_row(r, "reset")
            wait_for_empty_status_file()
            highlight_row(next_lesson[0], "#90EE90")  # highlight next lesson light green
            r = next_lesson[0]
        else:
            if r is not None:
                wait_for_empty_status_file()
                highlight_row(r, "reset")
                r = None
            print("No upcoming lessons.")

        if current_lesson:
            print(f"Current lesson is: {current_lesson[1]} at {excel_time_to_string(current_lesson[7])} to {excel_time_to_string(current_lesson[8])}")
            if s is not None and s != current_lesson[0]:
                wait_for_empty_status_file()
                highlight_row(s, "reset")
            wait_for_empty_status_file()
            highlight_row(current_lesson[0], "#EEE880")
            s = current_lesson[0]
        else:
            if s is not None:
                wait_for_empty_status_file()
                highlight_row(s, "reset")
                s = None
            print("No current lesson is in session.")

        time.sleep(reinit_delay)  # wait for the reinit_delay before reinitializing and rechecking the lessons

highlight()