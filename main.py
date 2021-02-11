import PySimpleGUI as sg
from openpyxl.utils.exceptions import InvalidFileException
from reporter.exceptions import UnsupportedNameLength, UnrecognisedType
from reporter import renessans, alyans, ingossrah

sg.theme("Purple")

layout = [[sg.Text('Страховая компания')],
          [sg.Combo(["Ренессанс", "Альянс", "Ингосстрах"], default_value="Ренессанс",  key="-INSURANCE-", enable_events=True)],
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
                if values["-INSURANCE-"] == "Ренессанс":
                    renessans.create_detach_report(files, values["-FOLDER-"], values["-FILENAME-"])
                elif values["-INSURANCE-"] == "Альянс":
                    alyans.create_detach_report(files, values["-FOLDER-"], values["-FILENAME-"])
                elif values["-INSURANCE-"] == "Ингосстрах":
                    ingossrah.create_detach_report(files, values["-FOLDER-"], values["-FILENAME-"])
                window['-OUTPUT-'].update('Файл создан', text_color="white")
        except (TypeError, ValueError):
            sg.popup("Ошибка", "Дата открепления не распознана")
        except InvalidFileException:
            sg.popup_error("Неверный формат файла.", "\nПоддерживаются только .xlsx,.xlsm,.xltx,.xltm,.xls")
        except UnrecognisedType as e:
            sg.popup_error(e.expression, e.message)
        except Exception as e:
            sg.popup_error("Неизвестная ошибка", "{0} {1}".format(str(e.__class__), e.args))
    elif event == 'Создать прикрепление':
        try:
            if values["-FILES-"] == "":
                window['-OUTPUT-'].update('Файлы не выбраны', text_color="OrangeRed4")
            elif values["-FOLDER-"] == "":
                window['-OUTPUT-'].update('Папка не выбрана', text_color="OrangeRed4")
            elif values["-FILENAME-"] == "":
                window['-OUTPUT-'].update("Имя файла не указано", text_color="OrangeRed4")
            else:
                files = values["-FILES-"].split(";")
                if values["-INSURANCE-"] == "Ренессанс":
                    renessans.create_attach_report(files, values["-FOLDER-"], values["-FILENAME-"])
                elif values["-INSURANCE-"] == "Альянс":
                    alyans.create_attach_report(files, values["-FOLDER-"], values["-FILENAME-"])
                elif values["-INSURANCE-"] == "Ингосстрах":
                    ingossrah.create_attach_report(files, values["-FOLDER-"], values["-FILENAME-"])
                window['-OUTPUT-'].update('Файл создан', text_color="white")
        except InvalidFileException:
            sg.popup_error("Неверный формат файла.", "\nПоддерживаются только .xlsx,.xlsm,.xltx,.xltm,.xls")
        except UnsupportedNameLength as e:
            sg.popup_error(e.expression, e.message)
        except UnrecognisedType as e:
            sg.popup_error(e.expression, e.message)
        except Exception as e:
            sg.popup_error("Неизвестная ошибка", "{0} {1}".format(str(e.__class__), e.args))
window.close()
