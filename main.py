import PySimpleGUI as sg
import os
import searcher
import writer

CONST_DEFAULT_ELEMENT_SIZE_X = 50
CONST_DEFAULT_ELEMENT_SIZE_Y = 50


def build_layout():
    layout = [[sg.Text("Selecciona la ruta raíz. Se buscarán ficheros *.xls y *.xlsx:")],
              [sg.Input(key='-FOLDER-', disabled=True), sg.FolderBrowse()],
              [sg.Checkbox("Buscar en subcarpetas", key="-CHK_SUBFOLDERS-")],
              [sg.Text("Define qué celda se buscará en los excel.")],
              [sg.InputText(key='-CELL-'), sg.Text()],
              [sg.Button('Ok'), sg.Button('Quit')]]
    return layout


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
            sg.popup(f'Carpeta seleccionada: {selected_folder}\n'
                     f'Buscar en subcarpetas: {search_subfolders}\n'
                     f'Celda elegida: {selected_cell}')
            excel_obtained_data = (searcher
                                   .search(folder=selected_folder, is_subfolders=search_subfolders, cell=selected_cell))
            sg.popup(f'Archivos Excel que contienen la celda especificada y su contenido:{excel_obtained_data}')
            writer.write_excel_file(excel_obtained_data)
            sg.popup(f'Excel creado en {os.getcwd()}')

    window.close()
