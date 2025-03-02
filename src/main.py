from config.main import EXCEL_FILE_DIR, STATUS_FILE_DIR, SAVE_LOCK_DIR, debug_mode, debounce_interval
from src.helper_functions.functions import autosave
autosave(EXCEL_FILE_DIR, STATUS_FILE_DIR, SAVE_LOCK_DIR, debug_mode, debounce_interval)