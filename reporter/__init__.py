from reporter.config import header
from openpyxl import Workbook
import os
from os import listdir
from os.path import isdir, isfile, join
from reporter.utils import fit_columns_width
from reporter.utils import open_xls_as_xlsx, open_csv_as_xlsx
from reporter.alyans import AlyansDetach, AlyansAttach
from reporter.renessans import RenessansAttach, RenessansDetach
from reporter.sogaz import SogazAttach, SogazDetach
from reporter.ingossrah import IngosstrahAttach, IngosstrahDetach
from reporter.maks import MaksAttach, MaksDetach
from reporter.rosgosstrah import RosgosstrahFabric, RosgosstrahReport, RosgosstrahDetach, RosgosstrahAttach
from reporter.vsk import VSKReport

from typing import Callable

alyans_folders = ["альянс"]
renessans_folders = ["рен", "ренессанс"]
ingosstrah_folders = ["ингосстрах", "ингос"]
sogaz_folders = ["согаз"]
maks_folders = ["макс"]
rosgosstrah_folders = ["росгосстрах", "россгосстрах"]
vsk_folders = ['вск']
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
            [self.add_vsk_attach(join(a, f)) for f in listdir(a) if isfile(join(a, f))]
        for d in detach:
            [self.add_vsk_detach(join(d, f)) for f in listdir(d) if isfile(join(d, f))]

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

    def add_vsk_attach(self, file):
        if VSKReport.is_reportable(file):
            self.reports.append(VSKReport(file))

    def add_vsk_detach(self, file):
        if VSKReport.is_reportable(file):
            self.reports.append(VSKReport(file))


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
            data += report.get_data()
        for r in data:
            ws.append(r)
        fit_columns_width(ws)
        wb.save(os.path.join(self.save_directory, self.filename + ".xlsx"))
        return os.path.join(self.save_directory, self.filename + ".xlsx")


if __name__ == '__main__':
    pass
