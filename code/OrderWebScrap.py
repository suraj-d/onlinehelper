import sys
from os import path, startfile
from subprocess import Popen

from PyQt5.QtWidgets import QFileDialog, QApplication

from code.CommanFunction import read_excel_sheet, create_text_file
from json import dumps
from PyQt5 import uic
from PyQt5.QtCore import QSettings

gui_file = "gui/createOrderSheet.ui"
Ui_order_sheet, baseclass = uic.loadUiType(gui_file)


class OrderScrapWindow(baseclass):

    def __init__(self):
        super(OrderScrapWindow, self).__init__()
        uic.loadUi(gui_file, self)

        # function value
        self.file_path = None
        self.portal_selected = None
        self.folder_path = None
        self.save_file = None
        self.sheet_name = "Sheet1"

        # Disable button
        self.browse_button.setDisabled(True)
        self.generate_sheet_button.setDisabled(True)
        self.open_folder_button.setDisabled(True)

        # button connection
        self.portal_combo_box.addItems(["Select Company", "Amazon", "Flipkart"])
        self.portal_combo_box.currentIndexChanged.connect(self.selectionchange)  # get value on selection change
        self.file_path_input.textChanged.connect(self.validate_gen_name)
        self.gen_name_input.textChanged.connect(self.validate_gen_name)

        self.browse_button.clicked.connect(self.browse_files)
        self.generate_sheet_button.clicked.connect(self.to_scrap_txt)
        self.open_folder_button.clicked.connect(self.open_folder)

        self.show()

    def selectionchange(self, i):
        """
        get combo box value on select
        """
        try:
            self.portal_selected = self.portal_combo_box.currentText()
            self.browse_button.setDisabled(False)
        except Exception as e:
            print(e)

    def browse_files(self):
        try:
            setting_last_file_path = QSettings("Order Helper", "LastSetting")  # path in regedit
            last_open_file_path = setting_last_file_path.value("order_sheet_path")  # get value

            allowed_file_format = "Excel File (*.xls *.xlsx *.csv);;All Files (*.*)"

            filepath = QFileDialog.getOpenFileName(self, 'Open Order Sheet File', last_open_file_path,
                                                   allowed_file_format)

            if filepath[0] != "":
                self.portal_combo_box.setDisabled(True)
                self.file_path = path.normpath(filepath[0])  # backslash in folder path text
                self.folder_path, file_name = path.split(self.file_path)
                self.file_path_input.setText(self.file_path)
                setting_last_file_path.setValue("order_sheet_path", self.folder_path)
        except Exception as e:
            print(e)

    def validate_gen_name(self):
        self.open_folder_button.setDisabled(True)
        if self.gen_name_input.text() and self.file_path_input.text():
            self.portal_combo_box.setDisabled(True)
            self.generate_sheet_button.setDisabled(False)
        elif not self.gen_name_input.text() and not self.file_path_input.text():
            self.portal_combo_box.setDisabled(False)
        else:
            self.generate_sheet_button.setDisabled(True)

    def to_scrap_txt(self):
        try:
            excel_data = read_excel_sheet(self.file_path, self.sheet_name)
            ws = excel_data.get('ws')
            start_row = excel_data.get('start_row')
            last_row = excel_data.get('last_row')
            print(f'Start Row: {start_row}'
                  f'last Row: {last_row}'
                  f'Portal: {self.portal_selected}')
            data_string = get_json_string(self.portal_selected, ws, start_row, last_row)
            self.save_file = create_text_file(data_string.get('web_scrap_string'), self.folder_path,
                                              f'{self.portal_selected} web scrap')
            print(self.save_file)
            self.open_folder_button.setDisabled(False)

            if self.open_check_box.isChecked():
                self.auto_open_file()
        except Exception as e:
            print(e)

    def open_folder(self):
        try:
            Popen(fr'explorer /select,{self.file_path}')
        except Exception as e:
            print(e)

    def auto_open_file(self):
        startfile(self.save_file)

    def closeEvent(self, event):
        try:
            pass
        except Exception as e:
            print(e)


# MAIN FUNCTIONS
def scrap_amazon():
    return {
        'head': r'{"_id":"amazon_price","startUrl":',
        'url': r'https://sellercentral.amazon.in/orders-v3/order/',
        'tail': r""","selectors":[{"delay":0,"id":"order_id","multiple":false,"parentSelectors":["_root"],"regex":"","selector":".a-span12 span.a-text-bold","type":"SelectorText"},{"delay":0,"id":"sku","multiple":false,"parentSelectors":["orderDetaiBox"],"regex":"","selector":"div.product-name-column-word-wrap-break-all:nth-of-type(3) div","type":"SelectorText"},{"delay":0,"id":"rate","multiple":false,"parentSelectors":["orderDetaiBox"],"regex":"","selector":"td.a-text-right span","type":"SelectorText"},{"delay":0,"id":"orderDetaiBox","multiple":true,"parentSelectors":["_root"],"selector":".a-spacing-large tbody tr","type":"SelectorElement"},{"delay":0,"id":"buyer_name","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"span div > span:nth-of-type(1)","type":"SelectorText"},{"delay":0,"id":"gst_state","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"span span:nth-of-type(5)","type":"SelectorText"}]}"""
    }


def scrap_flipkart():
    return {
        'head': r'{"_id":"fk_order_with_click_delete","startUrl":',
        'url': u"""https://seller.flipkart.com/index.html#dashboard/my-orders?serviceProfile=seller-fulfilled&shipmentType=easy-ship&orderState=shipments_delivered&orderItemId=""",
        'tail': r""","selectors":[{"clickElementSelector":"td.clickable","clickElementUniquenessType":"uniqueText","clickType":"clickOnce","delay":2000,"discardInitialElements":"do-not-discard","id":"click","multiple":false,"parentSelectors":["_root"],"selector":"td.clickable","type":"SelectorElementClick"},{"delay":0,"id":"name","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"div.styles__ItemWrapperVertical-sc-9qxfse-1:nth-of-type(2) div:nth-of-type(2)","type":"SelectorText"},{"delay":0,"id":"state","multiple":false,"parentSelectors":["_root"],"regex":"","selector":".styles__ItemWrapperVertical-sc-9qxfse-1 div:nth-of-type(6)","type":"SelectorText"},{"delay":0,"id":"box","multiple":true,"parentSelectors":["_root"],"selector":"div.styles__OrderItemContainer-sc-9qxfse-6","type":"SelectorElement"},{"delay":0,"id":"orderid","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div.styles__ItemDetails-sc-1bxatwx-6:nth-of-type(2) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"sku","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div.styles__ItemDetails-sc-1bxatwx-6:nth-of-type(1) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"rate","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div:nth-of-type(1) div:nth-of-type(6) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"qty","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div:nth-of-type(1) div.styles__PriceParams-sc-1bxatwx-11:nth-of-type(2) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"}]}"""

    }


def get_json_string(portal_name, ws, start_row, last_row):
    head = None
    mid = []
    url = None
    tail = None
    if portal_name.lower() == 'amazon':
        head = scrap_amazon().get("head")
        url = scrap_amazon().get("url")
        tail = scrap_amazon().get("tail")
    elif portal_name.lower() == 'flipkart':
        head = scrap_flipkart().get("head")
        url = scrap_flipkart().get("url")
        tail = scrap_flipkart().get("tail")

    for row in ws.iter_rows(values_only=True, min_row=start_row, max_row=last_row):
        mid.append(f'{url}{row[0]}')

    return {
        'web_scrap_string': f'{head}{dumps(mid)}{tail}'
    }


if __name__ == '__main__':
    app = QApplication(sys.argv)
    windows = OrderScrapWindow()
    sys.exit(app.exec_())
