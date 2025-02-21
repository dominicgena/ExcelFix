from datetime import datetime
import os


def create_file(prefix, root, ext):
    current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    filename = f"{prefix}_{current_time}.{ext}"
    file_path = os.path.join(root, filename)

    # Ensure the directory exists
    os.makedirs(root, exist_ok=True)

    # Create the file
    with open(file_path, 'w') as file:
        file.write('')  # Create an empty log file

    return file_path

