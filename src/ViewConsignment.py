import sys

from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QTableWidgetItem, QAbstractItemView, QAbstractScrollArea, QHeaderView

from src.CommanFunction import get_sql_connection, get_table_data
from src.AddConsignment import AddConsignmentWindow

if __name__ == "__main__":
    gui_file = "../gui/viewConsignment.ui"
else:
    gui_file = "gui/viewConsignment.ui"

Ui_order_form, baseClass = uic.loadUiType(gui_file)

# todo: add status check box to get consignment by status


class ViewConsignmentWindow(baseClass):

    def __init__(self):
        super(ViewConsignmentWindow, self).__init__()
        uic.loadUi(gui_file, self)

        # get data from sql
        self.load_consignment_data()
        self.load_status_data()

        # set table widget property and header
        self.sql_table.verticalHeader().setVisible(False)
        self.sql_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        header_list = ["SrNo.", "Consignment Id", "Pickup Date", "Status", "Warehouse Id", "Portal"]
        self.sql_table.setHorizontalHeaderLabels(header_list)
        self.sql_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # set column width
        header = self.sql_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        # header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.Stretch)

        self.sku_sql_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.sku_sql_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # Button
        self.new_button.clicked.connect(self.new_consignment)
        self.view_all_button.clicked.connect(self.show_hide_consignment)
        self.view_all_flag = False
        self.save_button.clicked.connect(self.save_consignment)
        self.refresh_button.clicked.connect(self.load_consignment_data)

        self.sql_table.itemClicked.connect(self.load_sku_table)
        self.sql_table.doubleClicked.connect(self.edit_consignment)

        # disable buttons
        self.save_button.setDisabled(True)
        # self.delete_button.setDisabled(True)

        self.show()

    @classmethod
    def new_consignment(cls):
        cls.window = AddConsignmentWindow()

    def edit_consignment(self, item):
        try:
            row = item.row()
            consignment_column = 1
            consignment_id = self.sql_table.item(row, consignment_column).text()

            status_column = 3
            status = self.sql_table.item(row, status_column).text()

            self.consignment_input.setText(consignment_id)
            self.status_input.setCurrentText(status)

            self.save_button.setDisabled(False)
        except Exception as e:
            print(e)

    def save_consignment(self):
        try:
            consignment_id = self.consignment_input.text()
            # warehouse_name = ""
            # create_date = ""
            # pickup_date = ""
            # delivery_date = ""
            status = self.status_input.currentText()

            print(f"{consignment_id}, {status}")

            update_consignment(consignment_id, status)
            self.load_consignment_data()
            self.save_button.setDisabled(True)
        except Exception as e:
            print(e)

    def load_status_data(self):
        query = "select status from warehouse_consignment_status"
        result = get_table_data(query).get('data')
        self.status_input.addItem("")
        for row_number, row_data in enumerate(result):
            for column_number, data in enumerate(row_data):
                self.status_input.addItem(data)

    def load_consignment_data(self):
        try:
            while self.sql_table.rowCount() > 0:
                self.sql_table.removeRow(0)
            # self.sql_table.clear()

            query = "SELECT id,consignment_id, pickup_date,status,warehouse,portal_name from " \
                    "view_warehouse_consignment where status <> 'Dispatched' and status<>'Delivered' order by " \
                    "pickup_date, consignment_id asc "

            result = get_table_data(query).get('data')

            # set in table widget
            self.sql_table.setRowCount(0)
            self.sql_table.setColumnCount(6)

            for row_number, row_data in enumerate(result):
                # print(row_number)
                self.sql_table.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    # print(column_number)
                    self.sql_table.setItem(row_number, column_number, QTableWidgetItem(str(data)))

            # self.sql_table.horizontalHeaderItem().setTextAlignment(Qt.AlignHCenter)

        except Exception as e:
            print(f"load_consignment_data: {e}")

    def load_sku_table(self, item):
        try:
            row = item.row()
            column = 1
            consignment_id = self.sql_table.item(row, column).text()

            # args = (consignment_id,)
            query = "select design_name,quantity from `view_warehouse_consignment_sku` where consignment_id =%s"
            result = get_table_data(query, (consignment_id,)).get('data')

            self.sku_sql_table.setRowCount(0)
            self.sku_sql_table.setColumnCount(2)
            for row_number, row_data in enumerate(result):
                # print(row_number)
                self.sku_sql_table.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    # print(column_number)
                    self.sku_sql_table.setItem(row_number, column_number, QTableWidgetItem(str(data)))
                    # print(data)

            header_list = [consignment_id, "Quantity"]
            self.sku_sql_table.setHorizontalHeaderLabels(header_list)
            header = self.sku_sql_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.Stretch)
            header.setSectionResizeMode(1, QHeaderView.ResizeToContents)

            self.consignment_input.clear()
            self.status_input.setCurrentIndex(0)
            self.save_button.setDisabled(True)
        except Exception as e:
            print(e)

    def show_hide_consignment(self):
        try:

            while self.sql_table.rowCount() > 0:
                self.sql_table.removeRow(0)
            # self.sql_table.clear()

            if self.view_all_flag:
                query = "SELECT id,consignment_id, pickup_date,status,warehouse,portal_name from " \
                        "view_warehouse_consignment where status <> 'Dispatched' and status <> 'Delivered' order by " \
                        "pickup_date, consignment_id asc "
                self.view_all_flag = False
                self.view_all_button.setText("View All")
            else:
                query = "SELECT id,consignment_id, pickup_date,status,warehouse,portal_name  from " \
                        "view_warehouse_consignment where status <> 'Delivered' " \
                        "order by pickup_date, consignment_id asc"
                self.view_all_flag = True
                self.view_all_button.setText("Hide")

            result = get_table_data(query).get('data')

            for row_number, row_data in enumerate(result):
                # print(row_number)
                self.sql_table.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    # print(column_number)
                    self.sql_table.setItem(row_number, column_number, QTableWidgetItem(str(data)))
        except Exception as e:
            print(f"show_hide_consignment: {e}")

#####################
### MAIN FUNCTION ###
#####################


def update_consignment(consignment_id, status):
    sql_validate = get_sql_connection()
    if 'error' in sql_validate:
        error = sql_validate.get("error")
        return {'error': error}

    mysqldb = sql_validate.get('mysqldb')
    cursor = mysqldb.cursor()

    args = [consignment_id, status]
    cursor.callproc("update_warehouse_consignment", args)

    mysqldb.commit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = ViewConsignmentWindow()

    sys.exit(app.exec_())
