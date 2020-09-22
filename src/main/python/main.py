from fbs_runtime.application_context.PyQt5 import ApplicationContext
from PyQt5.QtWidgets import QMainWindow, QLabel, QWidget, QVBoxLayout, QPushButton, QMessageBox

import sys
from ui.main_window import MainWindow



if __name__ == '__main__':

    appctxt = ApplicationContext()       # 1. Instantiate ApplicationContext

    window = MainWindow(appctxt)
    window.show()

    exit_code = appctxt.app.exec_()      # 2. Invoke appctxt.app.exec_()
    sys.exit(exit_code)