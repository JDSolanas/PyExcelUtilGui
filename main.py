import PySimpleGUI as sg

CONST_DEFAULT_ELEMENT_SIZE_X = 50
CONST_DEFAULT_ELEMENT_SIZE_Y = 50


def build_layout():
    layout = [[sg.Text("Ingresa un valor:")],
              [sg.Input(key='-FILE-', disabled=True), sg.FolderBrowse()],
              [sg.Button('Ok'), sg.Button('Quit')]]
    return layout


if __name__ == '__main__':
    window = sg.Window("PyExcelUtilGui", build_layout(), (CONST_DEFAULT_ELEMENT_SIZE_X, CONST_DEFAULT_ELEMENT_SIZE_Y))

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Quit'):
            break
        if event == 'Ok':
            selected_file = values['-FILE-']
            sg.popup(f'Ruta del archivo seleccionado: {selected_file}')

    window.close()
