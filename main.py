import PySimpleGUI as sg

CONST_DEFAULT_ELEMENT_SIZE_X = 50
CONST_DEFAULT_ELEMENT_SIZE_Y = 50


def build_layout():
    layout = [[sg.Text("Selecciona la ruta raíz. Se buscarán ficheros *.xls y *.xlsx:")],
              [sg.Input(key='-FOLDER-', disabled=True), sg.FolderBrowse()],
              [sg.Checkbox("Buscar en subcarpetas", key="-CHK_SUBFOLDERS-")],
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
            sg.popup(f'Carpeta seleccionada: {selected_folder}\nBuscar en subcarpetas: {search_subfolders}')

    window.close()
