from config.main import STATUS_FILE_DIR
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32com.client
import pythoncom
from pathlib import Path
import threading

class StatusFileEventHandler(FileSystemEventHandler):
    def __init__(self, excel_file_path, status_file_path, lock_file_path, debug_mode, debounce_interval):
        self.excel_file_path = excel_file_path
        self.status_file_path = status_file_path
        self.lock_file_path = Path(lock_file_path)
        self.debug_mode = debug_mode
        self.debounce_interval = debounce_interval
        self.debounce_timer = None

        # Initialize COM in the same thread
        pythoncom.CoInitialize()
        self.excel_app = win32com.client.Dispatch("Excel.Application")

        try:
            self.workbook = self.excel_app.Workbooks.Open(self.excel_file_path, ReadOnly=False)
            self.log_info(f"Workbook opened: {self.workbook.Name}")
        except Exception as e:
            self.log_error(f"Error opening workbook: {e}")

        print(f"Initialized event handler for: {self.status_file_path}")

    def is_excel_running(self):
        try:
            # Try accessing an attribute of the Excel application to check if it's still connected
            _ = self.excel_app.Workbooks.Count
            return True
        except Exception:
            return False


    def log_error(self, message):
        print(f"Error: {message}")


    def log_info(self, message):
        if self.debug_mode:
            print(message)


    def on_modified(self, event):
        if event.src_path == self.status_file_path:
            if self.debounce_timer:
                self.debounce_timer.cancel()
            self.debounce_timer = threading.Timer(self.debounce_interval, self.save_excel_file)
            self.debounce_timer.start()


    def save_excel_file(self):
        pythoncom.CoInitialize()  # Ensure COM is initialized in this thread

        if self.lock_file_path.exists():
            self.log_info("Save operation already in progress. Skipping...")
            return

        self.lock_file_path.touch()
        self.log_info(f"Lock file created: {self.lock_file_path}")

        try:
            # Ensure workbook is still connected
            if not self.is_excel_running():
                self.log_info("Excel instance lost. Reconnecting...")
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                self.workbook = self.excel_app.Workbooks.Open(self.excel_file_path, ReadOnly=False)

            self.log_info(f"Saving workbook: {self.workbook.Name}")
            self.workbook.Save()
            self.log_info("Workbook saved successfully.")
        except Exception as e:
            self.log_error(f"Error saving workbook: {e}")
        finally:
            time.sleep(1)
            if self.lock_file_path.exists():
                self.lock_file_path.unlink()
                self.log_info(f"Lock file removed: {self.lock_file_path}")


def autosave(excel_file_path, status_file_path, lock_file_path, debug_mode, debounce_interval):
    if not os.path.isfile(status_file_path):
        print(f"Error: {status_file_path} is not a file")
        return

    directory = os.path.dirname(status_file_path)
    print(f"Monitoring directory: {directory}")
    event_handler = StatusFileEventHandler(excel_file_path, status_file_path, lock_file_path, debug_mode, debounce_interval)
    observer = Observer()
    observer.schedule(event_handler, path=directory, recursive=False)
    observer.start()
    print(f"Started monitoring {status_file_path} for changes...")

    try:
        while True:
            time.sleep(1)  # Keep the script running
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


def cls():
    for i in range(100):
        print("\n")


def debug_log(message, debug_mode):
    """Prints message only if debug_mode is True."""
    if debug_mode:
        print(message)


def excel_time_to_string(excel_time):
    hours = int(excel_time * 24)
    minutes = round((excel_time * 24 * 60) % 60)  # Use round() instead of int()
    period = "AM" if hours < 12 else "PM"
    hours = hours % 12  # Convert to 12-hour format
    hours = 12 if hours == 0 else hours  # Adjust for 12 AM / 12 PM case
    return f"{hours}:{minutes:02d} {period}"


def get_valid_start_seconds(reinit_delay):
    delay_multiples = []
    current = 0.0

    # Populate the list with valid reinit_delay multiples up to 60 (not including 60)
    while current < 60 - reinit_delay:
        delay_multiples.append(current)
        current += reinit_delay

    # Mirror values (add the last value before mirroring to the rest), don't exceed 60
    n = delay_multiples[-1]
    for i in range(1, len(delay_multiples)):
        next_val = delay_multiples[i] + n
        if next_val < 60 - reinit_delay:
            delay_multiples.append(next_val)

    return delay_multiples

def wait_for_empty_status_file():
    while True:
        if os.path.exists(STATUS_FILE_DIR):
            with open(STATUS_FILE_DIR, 'r') as file:
                content = file.read().strip()
                if not content:
                    break
        time.sleep(1)



