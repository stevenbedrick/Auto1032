from PyQt5.QtCore import QThread, pyqtSignal, QCoreApplication, QUrl, QProcess, QDir
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QMessageBox, QFileDialog, QApplication, QFormLayout, \
    QLineEdit, QComboBox, QLabel, QDialog, QDialogButtonBox
from fbs_runtime.application_context.PyQt5 import ApplicationContext
from ui.threads import *
from excel.drawing import DRAWING_INPUT_RESOURCE_PATH, DRAWING_REL_INPUT_RESOURCE_PATH, \
    PRINTER_SETTINGS_RESOURCE_INPUT_PATH, SHEET_REL_TEMPLATE_RESOURCE_PATH, LOGO_RESOURCE_PATH
import sys


def show_in_finder(path: str):
    # I wonder if this will work?
    if sys.platform == 'win32':
        args_to_send = [
            "/select,",
            QDir.toNativeSeparators(path)
        ]
        QProcess.startDetached("explorer", args_to_send)
    elif sys.platform == 'darwin':
        args_to_send = [
            "-e",
            "tell application \"Finder\"",
            "-e",
            "activate",
            "-e",
            f"select posix file \"{path}\"",
            "-e",
            "end tell",
            "-e"
            "return"
        ]
        QProcess.execute("/usr/bin/osascript", args_to_send)


class DynamicDropdown(QComboBox):
    def __init__(self, list_entries=[], parent=None):
        super(DynamicDropdown, self).__init__(parent)

        for i in list_entries:
            self.addItem(i)

        if self.count() == 0:
            self.setEnabled(False)

        self.setSizeAdjustPolicy(QComboBox.AdjustToContents)

    def updateItems(self, new_items=[]):
        self.clear()
        for i in new_items:
            self.addItem(i)
        if self.count() == 0:
            self.setEnabled(False)
        else:
            self.setEnabled(True)

class CompletionDialog(QDialog):
    def __init__(self, *args, **kwargs):
        super(CompletionDialog, self).__init__(*args, **kwargs)

        self.setWindowTitle("Done!")
        self.setModal(True)

        buttons = QDialogButtonBox.Ok

        self.buttonBox = QDialogButtonBox(buttons)
        self.buttonBox.accepted.connect(self.accept)

        self.label = QLabel("Done!")

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.label)
        self.layout.addWidget(self.buttonBox)
        self.setLayout(self.layout)


class MainWindow(QWidget):
    batch_load_thread = None
    template_load_thread = None
    run_process_thread = None

    updated_sheets = pyqtSignal(object)
    sheet_struct = {}
    input_data_path = None
    template_workbook_path = None

    def __init__(self, cxt: ApplicationContext):
        super(MainWindow, self).__init__()

        self.app_context = cxt

        # an item for showing error messages:
        self.error_label = QLabel("")

        # a "go" button
        self.go_button = QPushButton("Go!")
        self.go_button.clicked.connect(self.handle_go_button)

        # the form will live in a widget with a formlayout, which we'll then stack with other widgets
        self.choose_template_widget = QWidget()
        self.choose_template_layout = QFormLayout()

        # we're going to need a pushbutton for choosing the input template:
        self.choose_input_button = QPushButton("Input DR568...")
        self.choose_input_label = QLineEdit()
        self.choose_input_label.setEnabled(False)
        self.choose_template_layout.addRow(self.choose_input_button, self.choose_input_label)

        # sheet dropdown
        self.sheet_dropdown = DynamicDropdown()
        self.sheet_label = "Choose sheet:"
        self.choose_template_layout.addRow(self.sheet_label, self.sheet_dropdown)

        # batch dropdown
        self.batch_dropdown = DynamicDropdown()
        self.batch_label = "Choose batch:"
        self.choose_template_layout.addRow(self.batch_label, self.batch_dropdown)

        # button for choosing the template to use:
        self.chooseTemplateButton = QPushButton("Template File:")
        self.chooseTemplateButton.setEnabled(False)
        self.chooseTemplateLabel = QLineEdit()
        self.chooseTemplateLabel.setEnabled(False)
        self.choose_template_layout.addRow(self.chooseTemplateButton, self.chooseTemplateLabel)

        # template sheet dropdown:
        self.chooseTemplateSheetDropdown = DynamicDropdown()
        self.chooseTemplateSheetLabel = "Template Sheet:"
        self.choose_template_layout.addRow(self.chooseTemplateSheetLabel, self.chooseTemplateSheetDropdown)

        # now, tell the form widget to use the form layout:
        self.choose_template_widget.setLayout(self.choose_template_layout)

        # and set up the wrapper:
        self.masterLayout = QVBoxLayout()

        self.masterLayout.addWidget(self.choose_template_widget)
        self.masterLayout.addWidget(self.go_button)
        self.masterLayout.addWidget(self.error_label)

        self.setLayout(self.masterLayout)

        # now do events
        self.choose_input_button.clicked.connect(self.chooseInputClick)
        self.chooseTemplateButton.clicked.connect(self.choose_template_click)
        self.updated_sheets.connect(self.handle_updated_sheets)
        self.sheet_dropdown.currentTextChanged.connect(self.handle_sheet_selection)

    def get_path(self, msg: str) -> str:
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(self, msg, "", "", options=options)
        return filename

    def on_sheets_and_batch_ready(self, data):
        self.choose_input_button.setEnabled(True)
        self.updated_sheets.emit(data)
        self.batch_load_thread = None
        self.error_label.setText("")
        self.chooseTemplateButton.setEnabled(True)

    def on_sheets_and_batch_err(self, e):
        self.choose_input_button.setEnabled(True)
        self.error_label.setText(e)
        self.batch_load_thread = None

    def on_template_sheets_ready(self, data):
        self.chooseTemplateSheetDropdown.updateItems(data)
        self.template_load_thread = None
        self.chooseTemplateButton.setEnabled(True)
        self.error_label.setText("")

    def on_template_sheets_err(self, e):
        self.chooseTemplateButton.setEnabled(True)
        self.error_label.setText(e)

    def chooseInputClick(self):
        new_path = self.get_path("Choose input DR568")
        if new_path:
            self.input_data_path = new_path
            self.choose_input_button.setEnabled(False)
            self.error_label.setText("Loading batch numbers...")
            self.choose_input_label.setText(self.input_data_path)
            self.batch_load_thread = LoadSheetsAndBatchThread(self.input_data_path)
            self.batch_load_thread.done.connect(self.on_sheets_and_batch_ready)
            self.batch_load_thread.err.connect(self.on_sheets_and_batch_err)
            self.batch_load_thread.start()
            # and clear out other fields:
            if not self.template_workbook_path:
                self.template_workbook_path = False
                self.chooseTemplateButton.setEnabled(False)
                self.chooseTemplateLabel.setText("")
                self.chooseTemplateSheetDropdown.updateItems([])

    def choose_template_click(self):
        new_path = self.get_path("Choose template...")
        if new_path:
            self.template_workbook_path = new_path
            self.chooseTemplateButton.setEnabled(False)
            self.error_label.setText("Loading template...")
            self.chooseTemplateLabel.setText(self.template_workbook_path)
            self.template_load_thread = LoadSheetsThread(self.template_workbook_path)
            self.template_load_thread.done.connect(self.on_template_sheets_ready)
            self.template_load_thread.start()

    def handle_updated_sheets(self, obj):
        """
        Called when we have updated sheet/batch data
        :param obj:
        :return:
        """
        self.sheet_struct = obj
        self.sheet_dropdown.updateItems(sorted(self.sheet_struct.keys()))
        if len(self.sheet_struct.keys()) > 0:
            first_sheet = sorted(self.sheet_struct.keys())[0]
            sorted_batches = sorted(self.sheet_struct[first_sheet])
            self.batch_dropdown.updateItems([str(b) for b in sorted_batches])  # it wants str not int

    def handle_sheet_selection(self):
        """
        Called when the input sheet selection dropdown changes
        :return:
        """
        try:
            new_batches = self.sheet_struct[self.sheet_dropdown.currentText()]
            self.batch_dropdown.updateItems([str(b) for b in sorted(new_batches)])
        except Exception as e:
            self.error_label.setText(str(e))

    def handle_go_button(self):

        if not self.input_data_path:
            self.error_label.setText("No input selected!")
            return
        if not self.template_workbook_path:
            self.error_label.setText("No template selected!")
            return

        # now check the various dropdowns:
        if self.sheet_dropdown.currentText() == "":
            self.error_label.setText("No input sheet selected!")
            return
        if self.batch_dropdown.currentText() == "":
            self.error_label.setText("No batch selected!")
            return
        if self.chooseTemplateSheetDropdown.currentText() == "":
            self.error_label.setText("No template sheet selected!")
            return

        # get an output path:
        # default filename:
        default_fname = f"1032 Batch {self.batch_dropdown.currentText()}.xlsx"
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getSaveFileName(self, "Output filename", default_fname, "", options=options)

        if filename:
            self.error_label.setText(f"Starting to generate file at {filename}")
            self.output_file_name = filename

            # let's do this! Figure out some resources
            drawing_input_path = self.app_context.get_resource(DRAWING_INPUT_RESOURCE_PATH)
            drawing_rel_path = self.app_context.get_resource(DRAWING_REL_INPUT_RESOURCE_PATH)
            printer_settings_path = self.app_context.get_resource(PRINTER_SETTINGS_RESOURCE_INPUT_PATH)
            sheet_rel_path = self.app_context.get_resource(SHEET_REL_TEMPLATE_RESOURCE_PATH)

            self.run_process_thread = Generate1032Thread(
                data_input_path=self.input_data_path,
                data_input_sheet=self.sheet_dropdown.currentText(),
                batch_number=int(self.batch_dropdown.currentText()),
                template_path=self.template_workbook_path,
                template_sheet=self.chooseTemplateSheetDropdown.currentText(),
                output_path=filename,
                drawing_input_path=self.app_context.get_resource(DRAWING_INPUT_RESOURCE_PATH),
                drawing_rel_input_path=self.app_context.get_resource(DRAWING_REL_INPUT_RESOURCE_PATH),
                printer_settings_input_path=self.app_context.get_resource(PRINTER_SETTINGS_RESOURCE_INPUT_PATH),
                sheet_rel_template_path=self.app_context.get_resource(SHEET_REL_TEMPLATE_RESOURCE_PATH),
                logo_input_path=self.app_context.get_resource(LOGO_RESOURCE_PATH)
            )
            self.run_process_thread.done.connect(self.handle_run_process_done)
            self.run_process_thread.status.connect(self.handle_run_process_status)
            self.go_button.setEnabled(False)
            self.run_process_thread.start()

    def handle_run_process_status(self, obj):
        step_n, total_n = obj
        self.error_label.setText(f"Done with step {step_n}/{total_n}")

    def handle_run_process_err(self, e):
        self.error_label.setText(e)
        self.go_button.setEnabled(True)

    def handle_run_process_done(self, res):
        self.error_label.setText("Done!")
        self.go_button.setEnabled(True)

        announcement = CompletionDialog()
        if announcement.exec_():
            if self.output_file_name:
                show_in_finder(self.output_file_name)
            QCoreApplication.quit()
            #QDesktopServices.openUrl(QUrl(f"file://{self.output_file_name}"))
        # QCoreApplication.quit()