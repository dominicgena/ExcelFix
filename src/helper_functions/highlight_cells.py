import xlwings as xw


def highlight_row(row_number, color):
    """
    Highlights specific columns in the given row with the specified color.

    Args:
        row_number (int): The row number to highlight.
        color (str): Hex color code (e.g., "#FF0000" for red) or "reset" to clear formatting.
    """
    wb = xw.books.active
    sheet = wb.sheets["Lesson Schedule"]

    # Columns to highlight
    columns = ["A", "B", "C", "L", "N", "P", "Q", "R", "S"]

    if color.lower() == "reset":
        for col in columns:
            sheet.range(f"{col}{row_number}").color = None  # Reset color
        return

    # Convert hex to RGB (Excel uses (R, G, B))
    if color.startswith("#"):
        color = color.lstrip("#")
    rgb = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))  # Convert hex to (R, G, B)

    # Apply color to specified columns
    for col in columns:
        sheet.range(f"{col}{row_number}").color = rgb

