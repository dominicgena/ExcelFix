import os

debug_mode = False
# Define the root directory of the project.
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_FILE_DIR = os.path.join(ROOT_DIR, 'Guitar-lessons.xlsm')
CONFIG_DIR = os.path.join(ROOT_DIR, 'config')
LOGS_ROOT_DIR = os.path.join(ROOT_DIR, 'logs')
SAVE_LOGS_DIR = os.path.join(LOGS_ROOT_DIR, 'save')
TRIGGERS_DIR = os.path.join(ROOT_DIR, 'triggers')
SAVE_LOCK_DIR = os.path.join(TRIGGERS_DIR, 'save.lock') # when the excel workbook saves, a temporary lock file is placed here.  This will allow the program to gracefully handle autosaves
SAVE_TRIGGER_DIR = os.path.join(TRIGGERS_DIR, 'autosave-trigger.txt')
STATUS_FILE_DIR = os.path.join(TRIGGERS_DIR, 'statusbar-state.txt')
debounce_interval = 5 # delay in seconds program should wait after an edit before saving.  Resets for every consecutive edit. Saving will only execute after this timer reches 0

