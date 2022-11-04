import sys
from os import getlogin

from PyQt5 import uic
from PyQt5.QtCore import QSettings
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication

from src.GoogleSheets import *
from src.OpenURL import OrderUrlDialog, open_order_url
from src.OrderEntry import OrderEntryWindow
from src.OrderSheet import OrderSheetWindow
from src.OrderWebScrap import OrderScrapWindow
from src.ReturnSheet import ReturnSheetWindow
from src.PaymentEntry import PaymentEntryWindow
from src.AddNewSku import AddSkuWindow
from src.ViewConsignment import ViewConsignmentWindow

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

        self.consignment_check_button.clicked.connect(self.consignment_check)

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

    def add_sku(cls):
        cls.sku_window = AddSkuWindow()

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
    def consignment_check(cls):
        cls.consignment_window = ViewConsignmentWindow()

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
    # todo: improve load speed
    # todo: FOR TALLY
    #  get data from tally
    #  get order id and return order id from tally
    #  direct entry to tally, no text file create or save to mysql
    #  get stock item from tally
    #  create job challan entry in tally
    #  create job receive entry in tally
    #  create stock transfer entry in tally


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("icon.ico"))

    w = MainWindow()

    sys.exit(app.exec_())
