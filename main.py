import PySimpleGUI as sg
from openpyxl.utils.exceptions import InvalidFileException
from reporter.exceptions import UnsupportedNameLength, UnrecognisedType, PupsAppException
from reporter import ReportChain, ReportGenerator
import subprocess

sg.theme("Purple")

layout = [[sg.Text('Папка со страховыми', font="Bold")],
          [sg.Input(), sg.FolderBrowse(key="path", button_text="Выбрать")],
          [sg.Text('Папка для сохранения (не обязательно)', font="Bold")],
          [sg.Input(), sg.FolderBrowse(key="-FOLDER-", button_text="Выбрать")],
          [sg.Text('Имя файла для сохранения', font="Bold")],
          [sg.Input(key="-FILENAME-", default_text="report")],
          [sg.Button("Создать"), sg.Button("Отмена"), sg.Button("Показать в Finder", visible=False)],
          [sg.Text(key='-OUTPUT-', size=(60, 2))]
          ]
window = sg.Window('PupsApp', layout)

saved_path = None

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == 'Отмена':
        break
    elif event == 'Создать':
        try:
            if len(values["path"]) == 0:
                window['-OUTPUT-'].update('Файлы не выбраны', text_color="OrangeRed4")
            elif values["-FILENAME-"] == "":
                window['-OUTPUT-'].update('Имя файла не выбрано', text_color="OrangeRed4")
            else:
                if values["-FOLDER-"] == "":
                    values["-FOLDER-"] = values["path"]
                chain = ReportChain(values["path"])
                generator = ReportGenerator(chain, values["-FOLDER-"], values["-FILENAME-"])
                saved_path = generator.generate()
                window["Показать в Finder"].update(visible=True)
                window['-OUTPUT-'].update('Файл создан', text_color="white")
        except (TypeError, ValueError):
            sg.popup("Ошибка", "Дата открепления не распознана")
        except InvalidFileException:
            sg.popup_error("Неверный формат файла.", "\nПоддерживаются только .xlsx,.xlsm,.xltx,.xltm,.xls")
        except UnrecognisedType as e:
            sg.popup_error(e.expression, e.message)
        except UnsupportedNameLength as e:
            sg.popup_error(e.expression, e.message)
        except PupsAppException as e:
            sg.popup_error(e.expression, e.message)
        except Exception as e:
            sg.popup_error("Неизвестная ошибка", "{0} {1}".format(str(e.__class__), e.args))
    elif event == "Показать в Finder":
        subprocess.call(["open", "-R", saved_path])
window.close()
