import PySimpleGUI as sg
from openpyxl.utils.exceptions import InvalidFileException
from reporter.exceptions import UnsupportedNameLength, UnrecognisedType, PupsAppException
from reporter import renessans, alyans, ingossrah, sogaz, ReportChain, create_reports

sg.theme("Purple")

layout = [[sg.Text('Ренессанс', font="Bold")],
          [sg.Text('Прикрепления', size=(12, 1)), sg.Input(), sg.FilesBrowse(key="RenAttach", button_text="Выбрать")],
          [sg.Text('Открепления', size=(12, 1)), sg.Input(), sg.FilesBrowse(key="RenDetach", button_text="Выбрать")],
          [sg.HSeparator()],
          [sg.Text('Альянс', font="Bold")],
          [sg.Text('Прикрепления', size=(12, 1)), sg.Input(), sg.FilesBrowse(key="AlAttach", button_text="Выбрать")],
          [sg.Text('Открепления', size=(12, 1)), sg.Input(), sg.FilesBrowse(key="AlDetach", button_text="Выбрать")],
          [sg.HSeparator()],
          [sg.Text('Ингосстрах', font="Bold")],
          [sg.Text('Прикрепления', size=(12, 1)), sg.Input(), sg.FilesBrowse(key="InAttach", button_text="Выбрать")],
          [sg.Text('Открепления', size=(12, 1)), sg.Input(), sg.FilesBrowse(key="InDetach", button_text="Выбрать")],
          [sg.HSeparator()],
          [sg.Text('Согаз', font="Bold")],
          [sg.Text('Прикрепления', size=(12, 1)), sg.Input(), sg.FilesBrowse(key="SoAttach", button_text="Выбрать")],
          [sg.Text('Открепления', size=(12, 1)), sg.Input(), sg.FilesBrowse(key="SoDetach", button_text="Выбрать")],
          [sg.HSeparator()],
          [sg.Text('Папка для сохранения', font="Bold")],
          [sg.Input(), sg.FolderBrowse(key="-FOLDER-", button_text="Выбрать")],
          [sg.Text('Имя файла для сохранения', font="Bold")],
          [sg.Input(key="-FILENAME-", default_text="report")],
          [sg.Button("Создать"), sg.Button("Отмена")],
          [sg.Text(key='-OUTPUT-', size=(60, 2))]
          ]
window = sg.Window('PupsApp', layout)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == 'Отмена':
        break
    elif event == 'Создать':
        try:
            report_chains = []
            all_files = \
                values["RenAttach"].split(";") + values["RenDetach"].split(";") + \
                values["AlAttach"].split(";") + values["AlDetach"].split(";") + \
                values["InAttach"].split(";") + values["InDetach"].split(";") + \
                values["SoAttach"].split(";") + values["SoDetach"].split(";")
            all_files = [file for file in all_files if file != '']
            if len(all_files) == 0:
                window['-OUTPUT-'].update('Файлы не выбраны', text_color="OrangeRed4")
            elif values["-FOLDER-"] == "":
                window['-OUTPUT-'].update('Папка не выбрана', text_color="OrangeRed4")
            elif values["-FILENAME-"] == "":
                window['-OUTPUT-'].update('Имя файла не выбрано', text_color="OrangeRed4")
            else:
                if values["RenDetach"] != "":
                    report_chains.append(ReportChain(renessans.create_detach_report, values["RenDetach"].split(";")))
                if values["RenAttach"] != "":
                    report_chains.append(ReportChain(renessans.create_attach_report, values["RenAttach"].split(";")))
                if values["AlDetach"] != "":
                    report_chains.append(ReportChain(alyans.create_detach_report, values["AlDetach"].split(";")))
                if values["AlAttach"] != "":
                    report_chains.append(ReportChain(alyans.create_attach_report, values["AlAttach"].split(";")))
                if values["InDetach"] != "":
                    report_chains.append(ReportChain(ingossrah.create_detach_report, values["InDetach"].split(";")))
                if values["InAttach"] != "":
                    report_chains.append(ReportChain(ingossrah.create_attach_report, values["InAttach"].split(";")))
                if values["SoDetach"] != "":
                    report_chains.append(ReportChain(sogaz.create_detach_report, values["SoDetach"].split(";")))
                if values["SoAttach"] != "":
                    report_chains.append(ReportChain(sogaz.create_attach_report, values["SoAttach"].split(";")))
                create_reports(report_chains, values["-FOLDER-"], values["-FILENAME-"])
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
window.close()
