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

        # self.return_data = ['rto', 'customer return', 'wrong customer return', 'wrong courier return', 'cancel']
        self.return_data = {'rto': 0,
                            'customer return': 0,
                            'wrong customer return': 0,
                            'wrong courier return': 0,
                            'cancel': 0}
        self.return_list = [['Tracking id', 'Status']]

        self.return_type_input.setDisabled(True)
        self.record_button.clicked.connect(self.record_next)
        self.save_return_sheet_button.clicked.connect(self.save_return_sheet)

        return_data_string = f"Rto: {self.return_data.get('rto')}\n" \
                             f"Customer Return: {self.return_data.get('customer return')}\n" \
                             f"Wrong Courier Return: {self.return_data.get('wrong courier return')}\n" \
                             f"Wrong Customer Return: {self.return_data.get('wrong customer return')}\n" \
                             f"Cancel: {self.return_data.get('cancel')}\n"
        self.return_data_output.setPlainText(return_data_string)
        self.show()

    def record_next(self):
        try:
            user_input = self.user_input.text()
            status = self.return_type_input.text()

            if user_input.lower() in self.return_data:
                self.return_type_input.setText(user_input)

            if status != "" and user_input.lower() not in self.return_data:
                # add to return list
                self.return_list.append([user_input, status])

                # add to return count
                return_count = self.return_data.get(status)
                self.return_data.update({status: return_count + 1})

            self.user_input.clear()

            # update text field
            return_data_string = f"Rto: {self.return_data.get('rto')}\n" \
                                 f"Customer Return: {self.return_data.get('customer return')}\n" \
                                 f"Wrong Courier Return: {self.return_data.get('wrong courier return')}\n" \
                                 f"Wrong Customer Return: {self.return_data.get('wrong customer return')}\n" \
                                 f"Cancel: {self.return_data.get('cancel')}\n"
            self.return_data_output.setPlainText(return_data_string)
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
