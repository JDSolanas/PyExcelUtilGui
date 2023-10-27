import os
import openpyxl as opx
import xlrd


def cell_name_to_indexes(cell) -> [int, int]:
    col = 0
    row = 0
    for char in cell:
        if char.isalpha():
            col = col * 26 + ord(char) - ord('A') + 1
        elif char.isnumeric():
            row = int(char)
    return row - 1, col - 1


def get_value_from_cell(file_path, cell) -> str:
    if file_path.endswith('.xlsx'):
        try:
            workbook = opx.load_workbook(file_path)
            sheet = workbook.active
            cell_value = sheet[cell].value
            return cell_value
        except Exception as e:
            print(f"Error al abrir {file_path}: {str(e)}")
            return ""
    elif file_path.endswith('.xls'):
        try:
            workbook = xlrd.open_workbook(file_path)
            sheet = workbook.sheet_by_index(0)
            row, col = cell_name_to_indexes(cell)
            cell_value = sheet.cell_value(row, col)
            return cell_value
        except Exception as e:
            print(f"Error al abrir{file_path}: {str(e)}")
            return ""


def search(folder, cell, is_subfolders) -> {}:
    excel_data = {}

    for root, _, files in os.walk(folder):
        for file in files:
            if file.endswith(('.xlsx', '.xls')):
                file_path = os.path.join(root, file)
                cell_value = get_value_from_cell(file_path, cell)
                if cell_value not in [None, ""]:
                    excel_data[file_path] = cell_value

        if not is_subfolders:
            break
    return excel_data
