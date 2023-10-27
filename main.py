import PySimpleGUI as sg
import re
import os
import searcher
import writer

CONST_DEFAULT_ELEMENT_SIZE_X = 50
CONST_DEFAULT_ELEMENT_SIZE_Y = 50
CONST_RE_CELDA = r"^[A-Za-z]+\d+$"


def build_layout() -> [[]]:
    layout = [[sg.Text("Selecciona la ruta raíz. Se buscarán ficheros *.xls y *.xlsx:")],
              [sg.Input(key='-FOLDER-', disabled=True), sg.FolderBrowse()],
              [sg.Checkbox("Buscar en subcarpetas", key="-CHK_SUBFOLDERS-")],
              [sg.InputText(key='-CELL-', size=(20, 1)), sg.Text("<- Define qué celda se buscará en los excel")],
              [sg.Button('Ok'), sg.Button('Quit')]]
    return layout


def valida_folder(value_folder) -> bool:
    is_folder_ok = value_folder != ""
    if not is_folder_ok:
        sg.popup(f'{value_folder} No es una carpeta válida')
    return is_folder_ok


def valida_cell(value_cell) -> bool:
    is_cell_ok = bool(re.match(CONST_RE_CELDA, value_cell))
    if not is_cell_ok:
        sg.popup(f'{value_cell} No es un formato de celda de Excel válido')
    return is_cell_ok


if __name__ == '__main__':
    window = sg.Window("PyExcelUtilGui", build_layout(), (CONST_DEFAULT_ELEMENT_SIZE_X, CONST_DEFAULT_ELEMENT_SIZE_Y))

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Quit'):
            break
        if event == 'Ok':
            selected_folder = values['-FOLDER-']
            search_subfolders = values['-CHK_SUBFOLDERS-']
            selected_cell = values['-CELL-']

            # Validaciones previas
            folder_ok = valida_folder(selected_folder)
            if not folder_ok:
                continue
            cell_ok = valida_cell(selected_cell)
            if not cell_ok:
                continue

            # Obtención de datos
            excel_obtained_data = (searcher
                                   .search(folder=selected_folder, is_subfolders=search_subfolders, cell=selected_cell))

            # Escritura de datos
            writer.write_excel_file(excel_obtained_data)
            sg.popup(f'Excel creado en {os.getcwd()}')

    window.close()
