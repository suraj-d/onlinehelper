import sys

from PyQt5 import uic
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtWidgets import QApplication, QFileDialog

from src.CommanFunction import create_xlsx_file

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
        self.return_data_count = {'rto': 0,
                                  'customer return': 0,
                                  'wrong customer return': 0,
                                  'wrong courier return': 0,
                                  'cancel': 0}
        self.return_data_list = [['Tracking id', 'Status']]
        self.return_type_scan = None
        self.file_saved = False

        self.return_type_input.setDisabled(True)
        self.record_button.clicked.connect(self.record_next)
        self.save_return_sheet_button.clicked.connect(self.save_return_sheet)

        return_data_string = f"Rto: 0\n" \
                             f"Customer Return: 0\n" \
                             f"Wrong Courier Return: 0\n" \
                             f"Wrong Customer Return: 0\n" \
                             f"Cancel: 0\n" \
                             f"Total Return: 0"
        self.return_data_output.setPlainText(return_data_string)
        self.show()

    def record_next(self):
        try:
            user_input = self.user_input.text()
            # return_type_scan = ""

            if user_input.lower() in self.return_data_count:
                self.return_type_input.setText(user_input)
                self.return_type_scan = user_input

            if self.return_type_scan is not None and user_input.lower() not in self.return_data_count:
                # add to return list
                self.return_data_list.append([user_input, self.return_type_scan])

                # add to return count
                # get count to return type scanned
                return_count = self.return_data_count.get(self.return_type_scan.lower())
                # update count fot scanned return by 1
                self.return_data_count.update({self.return_type_scan.lower(): return_count + 1})

            if user_input.lower == "save":
                self.save_return_sheet()

            self.user_input.clear()

            # update text field
            rto = self.return_data_count.get('rto')
            customer_return = self.return_data_count.get('customer return')
            wrong_courier_return = self.return_data_count.get('wrong courier return')
            wrong_customer_return = self.return_data_count.get('wrong customer return')
            cancel = self.return_data_count.get('cancel')
            total = rto + customer_return + wrong_courier_return + wrong_customer_return + cancel
            return_data_string = f"Rto: {rto}\n" \
                                 f"Customer Return: {customer_return}\n" \
                                 f"Wrong Courier Return: {wrong_courier_return}\n" \
                                 f"Wrong Customer Return: {wrong_customer_return}\n" \
                                 f"Cancel: {cancel}\n" \
                                 f"Total Return: {total}"
            self.return_data_output.setPlainText(return_data_string)
        except Exception as e:
            self.return_data_output.setPlainText(str(e))

    def save_return_sheet(self):
        try:
            setting_last_file_path = QSettings("Order Helper", "LastSetting")  # path in regedit

            last_open_file_path = setting_last_file_path.value("return_sheet_path")  # get value
            filepath = QFileDialog.getSaveFileName(self, "Save Return Scan Sheet", last_open_file_path,
                                                   "Excel Files (*.xlsx)")
            sheet_name = "Return Scan"  # sheet name and workbook name
            if filepath[0] != "":
                create_xlsx_file(self.return_data_list, filepath[0], sheet_name)
                setting_last_file_path.setValue("return_sheet_path", filepath[0])
                self.file_saved = True
        except Exception as e:
            print(e)

    def keyPressEvent(self, event):
        try:
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                self.record_next()
        except Exception as e:
            print(e)

    def closeEvent(self, event):
        try:
            for key, value in self.return_data_count.items():
                if value > 0 and not self.file_saved:
                    self.save_return_sheet()
        except Exception as e:
            print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    windows = ReturnSheetWindow()
    sys.exit(app.exec_())
