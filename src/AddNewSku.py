# import sys
#
# from PyQt5 import uic
# from PyQt5.QtGui import QIcon
# from PyQt5.QtWidgets import QApplication
# from src.CommanFunction import get_sql_connection
#
# if __name__ == "__main__":
#     gui_file = "../gui/addSku.ui"
# else:
#     gui_file = "gui/addSku.ui"
#
# Ui_order_form, baseClass = uic.loadUiType(gui_file)
#
#
# class AddSkuWindow(baseClass):
#
#     def __init__(self, *args, **kwargs):
#         super(AddSkuWindow, self).__init__(*args, **kwargs)
#         uic.loadUi(gui_file, self)
#
#         self.sku = None
#         self.color = None
#
#         self.show()
#
#     def to_mysql(self):
#         self.sku = self.sku_value.text()
#         self.color = self.color_value.text()
#
#         sql_validation = get_sql_connection()
#         if "error" not in sql_validation:  # check mysql connection
#             mysqldb = sql_validation.get('mysqldb')  # get sql database
#             sql_data = send_sku_to_mysql(self.sku, self.color, mysqldb)
#             if 'error' in sql_data:
#                 sql_msg = sql_data.get('error')
#             else:
#                 sql_msg = f"{sql_data.get('number')} of sku saved to database"
#         else:
#             sql_msg = sql_validation.get('error')
#
#         self.result_value.setText(str(sql_msg))
#
#
# def send_sku_to_mysql(sku, color, mysqldb):
#     cursor = mysqldb.cursor()
#     new_data_count = 0
#     # insert query
#     args = [sku, color]
#     cursor.callproc('insert_new_sku', args)
#     mysqldb.commit()
#     new_data_count += cursor.rowcount
#     print(new_data_count)
#
#     return {'number': new_data_count}
#
#
# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     windows = AddSkuWindow()
#     sys.exit(app.exec_())
