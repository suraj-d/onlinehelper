import sys
from os import getlogin

from PyQt5 import uic
from PyQt5.QtCore import QSettings
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication

from code.GoogleSheets import *
from code.OpenURL import OrderUrlDialog, open_order_url
from code.OrderEntry import OrderEntryWindow
from code.OrderSheet import OrderSheetWindow
from code.OrderWebScrap import OrderScrapWindow
from code.ReturnSheet import ReturnSheetWindow
from code.PaymentSheet import PaymentEntryWindow

gui_file = "gui/mainApplication.ui"
Ui_order_form, baseClass = uic.loadUiType(gui_file)


class MainWindow(baseClass):

    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        uic.loadUi(gui_file, self)
        self.setWindowIcon(QIcon('icon.ico'))

        # google sheet for admin only
        if getlogin() != "sunfashionLap":
            self.sheet_group.setVisible(False)

        self.button_connection()

        self.show()

    def button_connection(self):
        # tally xml buttons
        self.order_entry_button.clicked.connect(self.order_entry)
        self.return_entry_button.clicked.connect(self.order_entry)
        self.payment_entry_button.clicked.connect(self.payment_entry)

        # order check
        self.order_check_button.clicked.connect(self.order_check)
        self.edit_order_url.clicked.connect(self.order_url)

        self.create_order_sheet_button.clicked.connect(self.order_sheet)
        self.create_return_sheet_button.clicked.connect(self.return_sheet)

        # sheets
        self.costing_sheet_button.clicked.connect(self.costing_sheet)
        self.price_cal_button.clicked.connect(self.price_sheet)
        self.stock_sheet_button.clicked.connect(self.stock_sheet)
        self.scrap_order_button.clicked.connect(self.scrap_order)

        # server button
        self.server_shutdown_button.clicked.connect(self.server_shutdown)
        self.server_restart_button.clicked.connect(self.server_restart)

    @classmethod
    def order_entry(cls):
        cls.order_window = OrderEntryWindow()

    @classmethod
    def payment_entry(cls):
        cls.payment_window = PaymentEntryWindow()

    @staticmethod
    def costing_sheet():
        open_costing_sheet()

    @staticmethod
    def price_sheet():
        open_price_calculator()

    @staticmethod
    def stock_sheet():
        open_stock_sheet()

    @staticmethod
    def server_shutdown():
        shutdown_server()

    @staticmethod
    def server_restart():
        restart_server()

    @classmethod
    def order_check(cls):
        try:
            # get url_value from regedit setting
            cls.text_value = QSettings('Order Helper', 'LastSetting').value('url_value')
            urls = cls.text_value.split()
            open_order_url(urls)
        except Exception as e:
            print(e)

    @classmethod
    def order_url(cls):
        cls.edit_url_window = OrderUrlDialog()

    @classmethod
    def order_sheet(cls):
        cls.order_sheet_window = OrderSheetWindow()

    @classmethod
    def return_sheet(cls):
        cls.return_sheet_window = ReturnSheetWindow()

    @classmethod
    def scrap_order(cls):
        cls.order_scrap = OrderScrapWindow()

    def closeEvent(self, event):
        try:
            pass
        except Exception as e:
            print(e)

    # todo: threading
    # todo: paste to filtered row gui
    #  with start row and last row in excel data scan
    #  number of paste repeat
    # todo: update sku in mysql gui
    # todo: sync mysql sku list with google stock sheet
    # todo: return scan excel sheet
    #  if scan rto or customer change status for next scan
    # todo: order scrap separate gui
    #  with start and last row
    #  show generate text in textbox with copy button
    #  remove create text file
    # todo: improve load speed


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("icon.ico"))

    w = MainWindow()

    sys.exit(app.exec_())
