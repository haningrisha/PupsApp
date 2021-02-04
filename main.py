import PySimpleGUI as sg
from openpyxl.utils.exceptions import InvalidFileException
from reporter import create_detach_report, create_attach_report

sg.theme("Purple")

layout = [[sg.Text('Страховая компания')],
          [sg.Combo(["Ренесанс"], default_value="Ренесанс",  key="-INSURANCE-")],
          [sg.Text('Отчеты')],
          [sg.Input(), sg.FilesBrowse(key="-FILES-")],
          [sg.Text('Папка для сохранения')],
          [sg.Input(), sg.FolderBrowse(key="-FOLDER-")],
          [sg.Text('Имя файла для сохранения')],
          [sg.Input(key="-FILENAME-", default_text="report")],
          [sg.Button("Создать прикрепление"), sg.Button("Создать открепление"), sg.Button("Отмена")],
          [sg.Text(key='-OUTPUT-', size=(60, 2))]
          ]
window = sg.Window('PupsApp', layout)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == 'Отмена':
        break
    elif event == 'Создать открепление':
        try:
            if values["-FILES-"] == "":
                window['-OUTPUT-'].update('Файлы не выбраны', text_color="OrangeRed4")
            elif values["-FOLDER-"] == "":
                window['-OUTPUT-'].update('Папка не выбрана', text_color="OrangeRed4")
            elif values["-FILENAME-"] == "":
                window['-OUTPUT-'].update('Имя файла не выбрано', text_color="OrangeRed4")
            else:
                files = values["-FILES-"].split(";")
                create_detach_report(files, values["-FOLDER-"], values["-FILENAME-"])
                window['-OUTPUT-'].update('Файл создан', text_color="white")
        except (TypeError, ValueError):
            sg.popup("Ошибка", "Дата открепления не распознана")
        except InvalidFileException:
            sg.popup_error("Неверный формат файла.", "\nПоддерживаются только .xlsx,.xlsm,.xltx,.xltm,.xls")
        except:
            sg.popup_error("Ошибка", "Неизвестная ошибка")
    elif event == 'Создать прикрепление':
        try:
            if values["-FILES-"] == "":
                window['-OUTPUT-'].update('Файлы не выбраны', text_color="OrangeRed4")
            elif values["-FOLDER-"] == "":
                window['-OUTPUT-'].update('Папка не выбрана', text_color="OrangeRed4")
            elif values["-FILENAME-"] == "":
                files = values["-FILES-"].split(";")
                create_attach_report(files, values["-FOLDER-"], values["-FILENAME-"])
                window['-OUTPUT-'].update('Файл создан', text_color="white")
        except InvalidFileException:
            sg.popup_error("Неверный формат файла.", "\nПоддерживаются только .xlsx,.xlsm,.xltx,.xltm,.xls")
        except:
            sg.popup_error("Ошибка", "Неизвестная ошибка")
window.close()
