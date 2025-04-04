import time

FILE_PATH = r"D:\Projects\Python\ExcelFix\triggers\statusbar-state.txt"
CHECK_INTERVAL = 0.01  # 20 milliseconds
def read_file():
    try:
        with open(FILE_PATH, 'r') as f:
            return f.read().strip()
    except FileNotFoundError:
        return None

def monitor_file():
    last_content = None
    while True:
        content = read_file()
        if content != last_content:
            print(f"File state: '{content}'")
            last_content = content
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    monitor_file()
