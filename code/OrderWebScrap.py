import sys
from os import path, startfile
from subprocess import Popen

from PyQt5.QtWidgets import QFileDialog, QApplication
from pyperclip import copy

from code.CommanFunction import read_excel_sheet, create_text_file
from json import dumps
from PyQt5 import uic
from PyQt5.QtCore import QSettings

if __name__ == "__main__":
    gui_file = "../gui/createWebScrap.ui"
else:
    gui_file = "gui/createWebScrap.ui"
Ui_order_sheet, baseclass = uic.loadUiType(gui_file)


class OrderScrapWindow(baseclass):

    def __init__(self):
        super(OrderScrapWindow, self).__init__()
        uic.loadUi(gui_file, self)

        # function value
        self.sheet_name_input.setText("Sheet1")

        # Disable button
        self.browse_button.setDisabled(True)
        self.read_excel_button.setDisabled(True)
        self.genereate_code_button.setDisabled(True)

        # button connection
        combo_box_item = ["Select Company", "Amazon Order Detail", "Flipkart Order Detail",
                          "Flipkart Payment Detail"]
        self.scrap_type_combo_box.addItems(combo_box_item)
        self.scrap_type_combo_box.currentIndexChanged.connect(self.selectionchange)  # get value on selection change

        self.excel_sheet_input.textChanged.connect(self.selectionchange)
        self.sheet_name_input.textChanged.connect(self.selectionchange)
        self.start_row_input.textChanged.connect(self.selectionchange)
        self.last_row_input.textChanged.connect(self.selectionchange)

        self.browse_button.clicked.connect(self.browse_files)
        self.read_excel_button.clicked.connect(self.read_excel)
        self.genereate_code_button.clicked.connect(self.to_scrap_txt)
        self.copy_code_button.clicked.connect(self.copy_text)

        self.show()

    def selectionchange(self, i):
        """
        get combo box value on select
        """
        try:
            if self.scrap_type_combo_box.currentIndex() > 0:
                self.browse_button.setDisabled(False)
                if self.excel_sheet_input.text() is not None:
                    self.read_excel_button.setDisabled(False)
            else:
                self.browse_button.setDisabled(True)
                self.read_excel_button.setDisabled(True)
                self.genereate_code_button.setDisabled(True)

        except Exception as e:
            print(e)

    def browse_files(self):
        try:

            allowed_file_format = "Excel File (*.xls *.xlsx)"

            filepath = QFileDialog.getOpenFileName(self, 'Open Order Sheet File', "/",
                                                   allowed_file_format)

            if filepath[0] != "":
                self.scrap_type_combo_box.setDisabled(True)
                file_path = path.normpath(filepath[0])  # backslash in folder path text
                self.excel_sheet_input.setText(file_path)
        except Exception as e:
            print(e)

    def read_excel(self):
        try:
            excel_sheet_path = self.excel_sheet_input.text()
            sheet_name = self.sheet_name_input.text()
            excel_data = read_excel_sheet(excel_file_path=excel_sheet_path, sheet_name=sheet_name)
            if 'error' not in excel_data:
                self.start_row_input.setText(str(excel_data.get('start_row')))
                self.last_row_input.setText(str(excel_data.get('last_row')))
        except Exception as e:
            print(e)

    def to_scrap_txt(self):
        try:
            portal_selected = self.scrap_type_combo_box.currentText()
            excel_sheet_path = self.excel_sheet_input.text()
            sheet_name = self.sheet_name_input.text()
            excel_data = read_excel_sheet(excel_sheet_path, sheet_name)
            ws = excel_data.get('ws')

            start_row = int(self.start_row_input.text())
            last_row = int(self.last_row_input.text())

            data_string = get_web_scrap_json_string(portal_selected, ws, start_row, last_row)
            self.code_input.setPlainText(data_string.get('web_scrap_string'))

        except Exception as e:
            print(f'to_scrap_txt error: {e}')

    def copy_text(self):
        try:
            copy(self.code_input.toPlainText())
            self.copy_code_button.setText("Copied")
        except Exception as e:
            print(e)

    def closeEvent(self, event):
        try:
            pass
        except Exception as e:
            print(e)


# MAIN FUNCTIONS
# def scrap_amazon_order():
#     return {
#         'head': r'{"_id":"amazon_price","startUrl":',
#         'url': u'https://sellercentral.amazon.in/orders-v3/order/',
#         'tail': r',"selectors":[{"delay":0,"id":"order_id","multiple":false,"parentSelectors":["_root"],"regex":"","selector":".a-span12 span.a-text-bold","type":"SelectorText"},{"delay":0,"id":"sku","multiple":false,"parentSelectors":["orderDetaiBox"],"regex":"","selector":"div.product-name-column-word-wrap-break-all:nth-of-type(3) div","type":"SelectorText"},{"delay":0,"id":"rate","multiple":false,"parentSelectors":["orderDetaiBox"],"regex":"","selector":"td.a-text-right span","type":"SelectorText"},{"delay":0,"id":"orderDetaiBox","multiple":true,"parentSelectors":["_root"],"selector":".a-spacing-large tbody tr","type":"SelectorElement"},{"delay":0,"id":"buyer_name","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"span div > span:nth-of-type(1)","type":"SelectorText"},{"delay":0,"id":"gst_state","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"span span:nth-of-type(5)","type":"SelectorText"}]}'
#     }
#
#
# def scrap_flipkart_order():
#     return {
#         'head': r'{"_id":"fk_order_with_click_delete","startUrl":',
#         'url': u'https://seller.flipkart.com/index.html#dashboard/my-orders?serviceProfile=seller-fulfilled&shipmentType=easy-ship&orderState=shipments_delivered&orderItemId=',
#         'tail': r',"selectors":[{"clickElementSelector":"td.clickable","clickElementUniquenessType":"uniqueText","clickType":"clickOnce","delay":2000,"discardInitialElements":"do-not-discard","id":"click","multiple":false,"parentSelectors":["_root"],"selector":"td.clickable","type":"SelectorElementClick"},{"delay":0,"id":"name","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"div.styles__ItemWrapperVertical-sc-9qxfse-1:nth-of-type(2) div:nth-of-type(2)","type":"SelectorText"},{"delay":0,"id":"state","multiple":false,"parentSelectors":["_root"],"regex":"","selector":".styles__ItemWrapperVertical-sc-9qxfse-1 div:nth-of-type(6)","type":"SelectorText"},{"delay":0,"id":"box","multiple":true,"parentSelectors":["_root"],"selector":"div.styles__OrderItemContainer-sc-9qxfse-6","type":"SelectorElement"},{"delay":0,"id":"orderid","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div.styles__ItemDetails-sc-1bxatwx-6:nth-of-type(2) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"sku","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div.styles__ItemDetails-sc-1bxatwx-6:nth-of-type(1) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"rate","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div:nth-of-type(1) div:nth-of-type(6) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"qty","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div:nth-of-type(1) div.styles__PriceParams-sc-1bxatwx-11:nth-of-type(2) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"}]}'
#     }
#
#
# def scrap_flipkart_payment():
#     return {
#         'head': r'{"_id":"fk_payment_rec_rate","startUrl":',
#         'url': u'https://seller.flipkart.com/index.html#dashboard/payments/transactions?filter=',
#         'tail': r',"selectors":[{"delay":0,"id":"order_id","multiple":false,"parentSelectors":["_root"],"regex":"","selector":".cf-list td:nth-of-type(1)","type":"SelectorText"},{"delay":0,"id":"payment","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"#blinx-wrapper-38 span.total-amount","type":"SelectorText"}]}'
#     }


scrap_code = {'Amazon Order Detail': {
    'head': r'{"_id":"amazon_price","startUrl":',
    'url': u'https://sellercentral.amazon.in/orders-v3/order/',
    'tail': r',"selectors":[{"delay":0,"id":"order_id","multiple":false,"parentSelectors":["_root"],"regex":"","selector":".a-span12 span.a-text-bold","type":"SelectorText"},{"delay":0,"id":"sku","multiple":false,"parentSelectors":["orderDetaiBox"],"regex":"","selector":"div.product-name-column-word-wrap-break-all:nth-of-type(3) div","type":"SelectorText"},{"delay":0,"id":"rate","multiple":false,"parentSelectors":["orderDetaiBox"],"regex":"","selector":"td.a-text-right span","type":"SelectorText"},{"delay":0,"id":"orderDetaiBox","multiple":true,"parentSelectors":["_root"],"selector":".a-spacing-large tbody tr","type":"SelectorElement"},{"delay":0,"id":"buyer_name","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"span div > span:nth-of-type(1)","type":"SelectorText"},{"delay":0,"id":"gst_state","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"span span:nth-of-type(5)","type":"SelectorText"}]}'
},
    'Flipkart Order Detail': {
        'head': r'{"_id":"fk_order_with_click_delete","startUrl":',
        'url': u'https://seller.flipkart.com/index.html#dashboard/my-orders?serviceProfile=seller-fulfilled&shipmentType=easy-ship&orderState=shipments_delivered&orderItemId=',
        'tail': r',"selectors":[{"clickElementSelector":"td.clickable","clickElementUniquenessType":"uniqueText","clickType":"clickOnce","delay":2000,"discardInitialElements":"do-not-discard","id":"click","multiple":false,"parentSelectors":["_root"],"selector":"td.clickable","type":"SelectorElementClick"},{"delay":0,"id":"name","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"div.styles__ItemWrapperVertical-sc-9qxfse-1:nth-of-type(2) div:nth-of-type(2)","type":"SelectorText"},{"delay":0,"id":"state","multiple":false,"parentSelectors":["_root"],"regex":"","selector":".styles__ItemWrapperVertical-sc-9qxfse-1 div:nth-of-type(6)","type":"SelectorText"},{"delay":0,"id":"box","multiple":true,"parentSelectors":["_root"],"selector":"div.styles__OrderItemContainer-sc-9qxfse-6","type":"SelectorElement"},{"delay":0,"id":"orderid","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div.styles__ItemDetails-sc-1bxatwx-6:nth-of-type(2) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"sku","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div.styles__ItemDetails-sc-1bxatwx-6:nth-of-type(1) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"rate","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div:nth-of-type(1) div:nth-of-type(6) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"},{"delay":0,"id":"qty","multiple":false,"parentSelectors":["box"],"regex":"","selector":"div:nth-of-type(1) div.styles__PriceParams-sc-1bxatwx-11:nth-of-type(2) div.styles__Value-sc-1bxatwx-14","type":"SelectorText"}]}'
    },
    'Flipkart Payment Detail': {
        'head': r'{"_id":"fk_payment_rec_rate","startUrl":',
        'url': u'https://seller.flipkart.com/index.html#dashboard/payments/transactions?filter=',
        'tail': r',"selectors":[{"delay":0,"id":"order_id","multiple":false,"parentSelectors":["_root"],"regex":"","selector":".cf-list td:nth-of-type(1)","type":"SelectorText"},{"delay":0,"id":"payment","multiple":false,"parentSelectors":["_root"],"regex":"","selector":"#blinx-wrapper-38 span.total-amount","type":"SelectorText"}]}'
    }}


def get_web_scrap_json_string(portal_name, ws, start_row, last_row):
    """
    :param portal_name:
    :param ws:
    :param start_row:
    :param last_row:
    :return: dict {'web_scrap_string': json_string}
    """
    head = scrap_code.get(portal_name).get('head')
    url = scrap_code.get(portal_name).get('url')
    mid = []
    tail = scrap_code.get(portal_name).get('tail')

    try:
        for row in ws.iter_rows(values_only=True, min_row=start_row, max_row=last_row):
            mid.append(f'{url}{row[0]}')
    except Exception:
        print('Error found')
    return {
        'web_scrap_string': f'{head}{dumps(mid)}{tail}'
    }


if __name__ == '__main__':
    app = QApplication(sys.argv)
    windows = OrderScrapWindow()
    sys.exit(app.exec_())
