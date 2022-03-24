import sys

from time import sleep
from PyQt5 import QtWidgets as Qtw
from PyQt5 import QtCore as Qtc
from PyQt5 import uic
from webbrowser import open

if __name__ == "__main__":
    gui_file = "../gui/editOrderURL.ui"
else:
    gui_file = "gui/editOrderURL.ui"

Ui_edit_order_url, baseclass = uic.loadUiType(gui_file)


class OrderUrlDialog(baseclass):
    def __init__(self):
        super(OrderUrlDialog, self).__init__()
        uic.loadUi(gui_file, self)

        self.get_urls()

        self.text_value = self.setting_urls.value('url_value')
        self.url_text.setPlainText(self.text_value)

        self.show()

    @classmethod
    def get_urls(cls):
        # create order helper file in regedit if not available and set it to variable
        # computer\hkey_current_user\software\order helper\urls
        # Order Helper > main folder, LastSetting > sub folder, url_value > value tin Urls
        cls.setting_urls = Qtc.QSettings('Order Helper', 'LastSetting')

    def closeEvent(self, event):
        try:
            # update regedit value on window close
            self.setting_urls.setValue('url_value', self.url_text.toPlainText())
        except Exception as e:
            print(f'Error in Order url dialog close: {e}')


def open_order_url(urls):
    # open in default browser
    for url in urls:
        open(url)
        sleep(1)


if __name__ == '__main__':
    app = Qtw.QApplication(sys.argv)
    w = OrderUrlDialog()

    sys.exit(app.exec_())

