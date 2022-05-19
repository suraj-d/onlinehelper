import sys
from datetime import timedelta

from PyQt5 import uic, QtCore
from PyQt5.QtWidgets import QApplication

from src.CommanFunction import get_sql_connection, get_table_data

if __name__ == "__main__":
    gui_file = "../gui/addConsignment.ui"
else:
    gui_file = r"gui/addConsignment.ui"

Ui_order_form, baseClass = uic.loadUiType(gui_file)

# todo: add duplicate consignment check before save


class AddConsignmentWindow(baseClass):

    def __init__(self):
        super(AddConsignmentWindow, self).__init__()
        uic.loadUi(gui_file, self)
        try:
            # set calender popup
            self.created_date.setCalendarPopup(True)
            self.pickup_date.setCalendarPopup(True)
            self.delivery_date.setCalendarPopup(True)

            # set to current dates
            self.created_date.setDateTime(QtCore.QDateTime.currentDateTime())
            self.pickup_date.setDateTime(QtCore.QDateTime.currentDateTime())
            self.delivery_date.setDateTime(QtCore.QDateTime.currentDateTime())

            # load portal name
            self.load_portal_name()

            # buttons
            self.save_button.clicked.connect(self.save_consignment)
            self.new_button.clicked.connect(self.new_consignment)
            self.portal_combo_box.currentTextChanged.connect(self.load_warehouse_name)

            self.save_button.setDisabled(True)

            # self.sku_sql_table.itemChanged.connect(self.item_changed)

            self.sku_sql_table.setRowCount(50)
            self.sku_sql_table.setColumnCount(2)
            header_list = ["Sku", "Quantity"]
            self.sku_sql_table.setHorizontalHeaderLabels(header_list)
            self.sku_sql_table.doubleClicked.connect(self.enable_save_button)

        except Exception as e:
            print(e)

        self.show()

    def new_consignment(self):
        self.portal_combo_box.setCurrentIndex(0)
        self.consignment_value.setText("")
        self.warehouse_combo_box.clear()
        self.created_date.setDateTime(QtCore.QDateTime.currentDateTime())
        self.pickup_date.setDateTime(QtCore.QDateTime.currentDateTime())
        self.delivery_date.setDateTime(QtCore.QDateTime.currentDateTime())
        self.save_button.setDisabled(True)

    def enable_save_button(self):
        self.save_button.setDisabled(False)

    def save_consignment(self):
        try:
            consignment_id = self.consignment_value.text().strip()
            company = self.portal_combo_box.currentText()
            warehouse = self.warehouse_combo_box.currentText()
            created_date = self.created_date.date().toPyDate()
            pickup_date = self.pickup_date.date().toPyDate()
            delivery_date = self.delivery_date.date().toPyDate()
            status = "Scheduled"

            consignment_update = insert_consignment(consignment_id, warehouse, created_date, pickup_date, delivery_date, status)

            sku_row_count = self.sku_sql_table.rowCount()
            # print('sku start')
            sku_update = 0
            for i in range(sku_row_count):
                if self.sku_sql_table.item(i, 0) is None:
                    break

                sku = self.sku_sql_table.item(i, 0).text().strip()
                qty = self.sku_sql_table.item(i, 1).text()

                sku_update += insert_sku(consignment_id, sku, qty)
                # print(f"{sku}\n"
                #       f"{qty}\n"
                #       f"--------")
            update = f"{consignment_update} consignment added, with {sku_update} sku"
            self.status_label.setText(update)
            self.save_button.setDisabled(True)
        except Exception as e:
            self.status_label.setText(str(e))
            print(f"save_consignment: {e}")

    def load_portal_name(self):
        self.portal_combo_box.addItem("Select Portal")

        query = "select name from portal_name"
        result = get_table_data(query).get("data")

        for i in result:
            self.portal_combo_box.addItems(i)

    def load_warehouse_name(self):
        try:
            if self.portal_combo_box.currentIndex() == 0:
                return
            portal_name = self.portal_combo_box.currentText()

            query = "select portal_name_id from portal_name where name = %s"
            portal_name_id = get_table_data(query, (portal_name,)).get('data')[0][0]

            query = "SELECT code From warehouse where portal_name_id = %s"
            result = get_table_data(query, (portal_name_id,)).get('data')
            # warehouse_name = get_warehouse_name(portal_name).get('warehouse')
            self.warehouse_combo_box.clear()
            self.warehouse_combo_box.addItem("Select Warehouse")
            for i in result:
                self.warehouse_combo_box.addItems(i)
        except Exception as e:
            print(f"load_warehouse_name_cb: {e}")


###################
###MAIN FUNCTION###
###################

def insert_consignment(consignment_id, warehouse_name, create_date, pickup_date, delivery_date, status):
    sql_validate = get_sql_connection()
    if 'error' in sql_validate:
        error = sql_validate.get("error")
        return {'error': error}

    mysqldb = sql_validate.get('mysqldb')
    cursor = mysqldb.cursor()

    args = [consignment_id, warehouse_name, create_date, pickup_date, delivery_date, status]
    cursor.callproc("insert_warehouse_consignment", args)

    mysqldb.commit()
    count = int(cursor.rowcount)
    return count


def insert_sku(consignment_id, sku_name, quantity):
    sql_validate = get_sql_connection()
    if 'error' in sql_validate:
        error = sql_validate.get("error")
        return {'error': error}

    mysqldb = sql_validate.get('mysqldb')
    cursor = mysqldb.cursor()

    args = [consignment_id, sku_name, quantity]
    cursor.callproc("insert_warehouse_consignment_sku", args)

    mysqldb.commit()
    count = int(cursor.rowcount)
    return count


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = AddConsignmentWindow()

    sys.exit(app.exec_())
