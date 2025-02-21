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

    def log_info(self, message):
        if self.debug_mode:
            print(message)

    def log_error(self, message):
        print(f"Error: {message}")

    def is_excel_running(self):
        try:
            # Try accessing an attribute of the Excel application to check if it's still connected
            _ = self.excel_app.Workbooks.Count
            return True
        except Exception:
            return False

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

    def on_modified(self, event):
        if event.src_path == self.status_file_path:
            if self.debounce_timer:
                self.debounce_timer.cancel()
            self.debounce_timer = threading.Timer(self.debounce_interval, self.save_excel_file)
            self.debounce_timer.start()

def monitor_status_file(excel_file_path, status_file_path, lock_file_path, debug_mode, debounce_interval):
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