import sys

from PyQt5 import uic
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QFileDialog
from openpyxl import Workbook

from code.CommanFunction import create_xlsx_file

if __name__ == "__main__":
    gui_file = "../gui/createReturnSheet.ui"
else:
    gui_file = "gui/createReturnSheet.ui"

Ui_order_form, baseClass = uic.loadUiType(gui_file)


class ReturnSheetWindow(baseClass):

    def __init__(self):
        super(ReturnSheetWindow, self).__init__()
        uic.loadUi(gui_file, self)

        self.return_type = ['rto', 'customer return', 'wrong customer return', 'wrong courier return', 'cancel']
        self.return_list = [['Tracking id', 'Status']]

        self.return_type_input.setDisabled(True)
        self.record_button.clicked.connect(self.record_next)
        self.save_return_sheet_button.clicked.connect(self.save_return_sheet)
        self.show()

    def record_next(self):
        try:
            user_input = self.user_input.text()
            status = self.return_type_input.text()

            if user_input.lower() in self.return_type:
                self.return_type_input.setText(user_input)

            if status != "" and user_input.lower() not in self.return_type:
                self.return_list.append([user_input, status])
                self.recorded_detail()
                print(self.return_list)

            self.user_input.clear()
        except Exception as e:
            print(e)

    def recorded_detail(self):
        try:
            courier_return= 0
            customer_return = 0
            wrong_courier_return = 0
            wrong_customer_return = 0
            cancel = 0

            print(self.return_list)


        except Exception as e:
            print(e)

    def save_return_sheet(self):
        filepath = QFileDialog.getSaveFileName(self, "Save Return Scan Sheet", "",
                                               "Excel Files (*.xlsx)")
        sheet_name = "Return Scan"
        create_xlsx_file(self.return_list, filepath[0], sheet_name)

    def keyPressEvent(self, event):
        try:
            if event.key() == Qt.Key_Return:
                self.record_next()
        except Exception as e:
            print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    windows = ReturnSheetWindow()
    sys.exit(app.exec_())
