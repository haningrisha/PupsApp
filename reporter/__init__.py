from .config import header
from openpyxl import Workbook
import os
from os import listdir
from os.path import isdir, isfile, join
from .utils import fit_columns_width
from .utils import open_xls_as_xlsx, open_csv_as_xlsx
from .alyans import AlyansDetach, AlyansAttach
from .renessans import RenessansAttach, RenessansDetach
from .sogaz import SogazAttach, SogazDetach
from .ingossrah import IngosstrahAttach, IngosstrahDetach
from .maks import MaksAttach, MaksDetach
from .rosgosstrah import RosgosstrahFabric, RosgosstrahReport, RosgosstrahDetach, RosgosstrahAttach
from .vsk import get_vsk
from .alfa import get_alfa
from .absolut import get_absolut
from .best_doctor import get_best_doctor

from typing import Callable

alyans_folders = ["альянс"]
renessans_folders = ["рен", "ренессанс"]
ingosstrah_folders = ["ингосстрах", "ингос"]
sogaz_folders = ["согаз"]
maks_folders = ["макс"]
rosgosstrah_folders = ["росгосстрах", "россгосстрах"]
vsk_folders = ['вск']
alfa_folders = ['альфа']
absolut_folders = ['абсолют']
best_doctor_folders = ['бест доктор', 'бестдоктор', 'best doctor']
detach_folders = ["откреп", "откр", "открепление", "открепления", "detach", "detachment", "detachments"]
attach_folders = ["прикреп", "прикр", "прикрепление", "прикрепления", "attach", "attachment", "attachments"]


class ReportChain:
    def __init__(self, folder):
        self.reports = []
        self.folder = folder
        self.add_from_folder()

    def add_from_folder(self):
        self.add_insurance(self.add_alyans, alyans_folders)
        self.add_insurance(self.add_renessans, renessans_folders)
        self.add_insurance(self.add_ingosstrah, ingosstrah_folders)
        self.add_insurance(self.add_sogaz, sogaz_folders)
        self.add_insurance(self.add_maks, maks_folders)
        self.add_insurance(self.add_rosgosstrah, rosgosstrah_folders)
        self.add_insurance(self.add_vsk, vsk_folders)
        self.add_insurance(self.add_alfa, alfa_folders)
        self.add_insurance(self.add_absolut, absolut_folders)
        self.add_insurance(self.add_best_doctor, best_doctor_folders)

    def add_insurance(self, adder: Callable, folder_names: list):
        only_dirs = [f for f in listdir(self.folder) if isdir(join(self.folder, f))]
        [adder(join(self.folder, f)) for f in only_dirs if f.lower() in folder_names]

    def add_alyans(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_alyans_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_alyans_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_renessans(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_renessans_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_renessans_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_ingosstrah(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_ingosstrah_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_ingosstrah_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_sogaz(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_sogaz_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_sogaz_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_maks(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_maks_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_maks_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_rosgosstrah(self, folder):
        only_files = [join(folder, file) for file in listdir(folder) if isfile(join(folder, file))]
        for file in only_files:
            if RosgosstrahReport.is_reportable(file):
                self.reports.append(RosgosstrahFabric.create_report(file))
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            [self.add_rosgosstrah_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_rosgosstrah_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

    def add_vsk(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            self.reports.extend([get_vsk(join(a, f), attach=True) for f in listdir(a) if isfile(join(a, f))])
        for d in detach:
            self.reports.extend([get_vsk(join(d, f), detach=True) for f in listdir(d) if isfile(join(d, f))])

    def add_alfa(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            self.reports.extend([get_alfa(join(a, f), attach=True) for f in listdir(a) if isfile(join(a, f))])
        for d in detach:
            self.reports.extend([get_alfa(join(d, f), detach=True) for f in listdir(d) if isfile(join(d, f))])

    def add_absolut(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            self.reports.extend([get_absolut(join(a, f), attach=True) for f in listdir(a) if isfile(join(a, f))])
        for d in detach:
            self.reports.extend([get_absolut(join(d, f), detach=True) for f in listdir(d) if isfile(join(d, f))])

    def add_best_doctor(self, folder):
        detach = self.get_detach(folder)
        attach = self.get_attach(folder)
        for a in attach:
            self.reports.extend([get_best_doctor(join(a, f), attach=True) for f in listdir(a) if isfile(join(a, f))])
        for d in detach:
            self.reports.extend([get_best_doctor(join(d, f), detach=True) for f in listdir(d) if isfile(join(d, f))])

    def get_detach(self, folder):
        only_dirs = [f for f in listdir(folder) if isdir(join(folder, f))]
        return [join(folder, f) for f in only_dirs if f.lower() in detach_folders]

    def get_attach(self, folder):
        only_dirs = [f for f in listdir(folder) if isdir(join(folder, f))]
        return [join(folder, f) for f in only_dirs if f.lower() in attach_folders]

    def add_alyans_attach(self, file):
        if AlyansAttach.is_reportable(file):
            self.reports.append(AlyansAttach(file))

    def add_alyans_detach(self, file):
        if AlyansDetach.is_reportable(file):
            self.reports.append(AlyansDetach(file))

    def add_sogaz_attach(self, file):
        if SogazAttach.is_reportable(file):
            self.reports.append(SogazAttach(file))

    def add_sogaz_detach(self, file):
        if SogazDetach.is_reportable(file):
            self.reports.append(SogazDetach(file))

    def add_ingosstrah_attach(self, file):
        if IngosstrahAttach.is_reportable(file):
            self.reports.append(IngosstrahAttach(file))

    def add_ingosstrah_detach(self, file):
        if IngosstrahDetach.is_reportable(file):
            self.reports.append(IngosstrahDetach(file))

    def add_renessans_attach(self, file):
        if RenessansAttach.is_reportable(file):
            self.reports.append(RenessansAttach(file))

    def add_renessans_detach(self, file):
        if RenessansDetach.is_reportable(file):
            self.reports.append(RenessansDetach(file))

    def add_maks_attach(self, file):
        if MaksAttach.is_reportable(file):
            self.reports.append(MaksAttach(file))

    def add_maks_detach(self, file):
        if MaksDetach.is_reportable(file):
            self.reports.append(MaksDetach(file))

    def add_rosgosstrah_attach(self, file):
        if RosgosstrahAttach.is_reportable(file):
            self.reports.append(RosgosstrahAttach(file))

    def add_rosgosstrah_detach(self, file):
        if RosgosstrahDetach.is_reportable(file):
            self.reports.append(RosgosstrahDetach(file))


class ReportGenerator:
    def __init__(self, report_chain, save_directory, filename):
        self.report_chain = report_chain
        self.save_directory = save_directory
        self.filename = filename

    def generate(self):
        data = [header]
        wb = Workbook()
        ws = wb.active
        for report in self.report_chain.reports:
            try:
                data += report.get_data()
            except Exception as e:
                if hasattr(e, 'message'):
                    e.message = e.message + f' в файле {report.file}'
                raise
        for r in data:
            ws.append(r)
        fit_columns_width(ws)
        wb.save(os.path.join(self.save_directory, self.filename + ".xlsx"))
        return os.path.join(self.save_directory, self.filename + ".xlsx")


if __name__ == '__main__':
    pass
