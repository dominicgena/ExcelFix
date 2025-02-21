import os
from src.helper_functions.functions import monitor_status_file

debug_mode = True
# Define the root directory of the project.
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_FILE_DIR = os.path.join(ROOT_DIR, 'Guitar-lessons.xlsm')
CONFIG_DIR = os.path.join(ROOT_DIR, 'config')
LOGS_ROOT_DIR = os.path.join(ROOT_DIR, 'logs')
SAVE_LOGS_DIR = os.path.join(LOGS_ROOT_DIR, 'save')
TRIGGERS_DIR = os.path.join(ROOT_DIR, 'triggers')
SAVE_LOCK_DIR = os.path.join(TRIGGERS_DIR, 'save.lock')
SAVE_TRIGGER_DIR = os.path.join(TRIGGERS_DIR, 'autosave-trigger.txt')
STATUS_FILE_DIR = os.path.join(TRIGGERS_DIR, 'statusbar-state.txt')
debounce_interval = 5

monitor_status_file(EXCEL_FILE_DIR, STATUS_FILE_DIR, SAVE_LOCK_DIR, debug_mode, debounce_interval)