from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QMessageBox, QFileDialog, QApplication, QFormLayout, \
    QLineEdit, QComboBox, QLabel
from data.loader import load_batch_numbers_from_inventory_file, load_sheet_names
from excel.driver import run_complete_process


class LoadSheetsThread(QThread):
    done = pyqtSignal(object)
    err = pyqtSignal(str)
    def __init__(self, path_to_spreadsheet):
        QThread.__init__(self)
        self.path = path_to_spreadsheet
    def run(self):
        try:
            sheet_names = load_sheet_names(self.path)
            self.done.emit(sheet_names)
        except Exception as e:
            self.err.emit(str(e))

class LoadSheetsAndBatchThread(QThread):
    done = pyqtSignal(object)
    err = pyqtSignal(str)

    def __init__(self, path_to_spreadsheet):
        QThread.__init__(self)
        self.path = path_to_spreadsheet

    def run(self):
        try:
            sheet_names = load_sheet_names(self.path)

            to_ret = {}
            for s in sheet_names:
                this_sheet_batch = load_batch_numbers_from_inventory_file(self.path, s)
                to_ret[s] = sorted([str(b) for b in this_sheet_batch])

            self.done.emit(to_ret)

        except Exception as e:
            self.err.emit(str(e))


class Generate1032Thread(QThread):
    status = pyqtSignal(object)
    done = pyqtSignal(object)
    err = pyqtSignal(str)

    def __init__(self, data_input_path: str,
        data_input_sheet: str,
        batch_number: str,
        template_path: str,
        template_sheet: str,
        output_path: str,
        drawing_input_path: str,
        drawing_rel_input_path: str,
        printer_settings_input_path: str,
        sheet_rel_template_path: str,
        logo_input_path: str
):
        QThread.__init__(self)
        self.data_input_path = data_input_path
        self.data_input_sheet = data_input_sheet
        self.batch_number = batch_number
        self.template_path = template_path
        self.template_sheet = template_sheet
        self.output_path = output_path
        self.drawing_input_path = drawing_input_path
        self.drawing_rel_input_path = drawing_rel_input_path
        self.printer_settings_input_path = printer_settings_input_path
        self.sheet_rel_template_path = sheet_rel_template_path
        self.logo_input_path = logo_input_path

    def run(self):
        print("Starting...")
        def cb(step_n, total_n):
            self.status.emit((step_n, total_n))

        run_complete_process(
            data_input_path=self.data_input_path,
            data_input_sheet=self.data_input_sheet,
            batch_number=self.batch_number,
            template_path=self.template_path,
            template_sheet=self.template_sheet,
            output_path=self.output_path,
            drawing_input_path=self.drawing_input_path,
            drawing_rel_input_path=self.drawing_rel_input_path,
            printer_settings_input_path=self.printer_settings_input_path,
            sheet_rel_template_path=self.sheet_rel_template_path,
            logo_input_path=self.logo_input_path,
            progress_callback=cb
        )

        self.done.emit("Done!")