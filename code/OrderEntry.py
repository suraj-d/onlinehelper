import sys
from datetime import datetime
from os import path
from pyperclip import copy

from PyQt5 import uic
from PyQt5.QtCore import QSettings
from PyQt5.QtGui import QIntValidator
from PyQt5.QtWidgets import QApplication, QFileDialog
from mysql.connector import connect

from code.xmlFormats import head_xml, order_body_xml, return_body_xml, tail_xml
from code.CommanFunction import read_excel_sheet, create_text_file

if __name__ == "__main__":
    gui_file = "../gui/orderToXML.ui"
else:
    gui_file = "gui/orderToXML.ui"

Ui_order_form, baseClass = uic.loadUiType(gui_file)


class OrderEntryWindow(baseClass):

    def __init__(self, *args, **kwargs):
        super(OrderEntryWindow, self).__init__(*args, **kwargs)
        uic.loadUi(gui_file, self)

        # get call button name
        if self.sender() is not None:
            self.sending_button = self.sender().objectName()
        else:
            self.sending_button = "order_entry_button"

        # ui call button names
        self.order_button = "order_entry_button"
        self.return_button = "return_entry_button"

        self.update_window_text()

        # set text value
        self.sheet_name_input.setText("For Tally")

        # function values
        # self.file_path = None
        # self.folder_path = None
        # self.file_name = None
        self.ws = None
        self.start_row = None
        self.last_row = None
        self.data_start_row = None
        self.data_last_row = None

        # only int for start and last row validation
        self.onlyInt = QIntValidator()
        self.last_row_input.setValidator(self.onlyInt)
        self.start_row_input.setValidator(self.onlyInt)

        self.button_on_load()
        self.button_connection()

        self.show()

    def update_window_text(self):
        # set window title and buttons name
        window_title = None
        to_xml_button_text = None
        to_sql_button_text = None
        excel_file_lbl = None
        if self.sending_button == self.order_button:
            window_title = "For Tally Order Entry"
            to_xml_button_text = "Order To Tally XML"
            to_sql_button_text = "Order To My SQL"
            excel_file_lbl = "Order Excel Sheet:"
        elif self.sending_button == self.return_button:
            window_title = "For Tally Return Entry"
            to_xml_button_text = "Return To Tally XML"
            to_sql_button_text = "Return To My SQL"
            excel_file_lbl = "Return Excel Sheet:"

        # set text as per call button clicked
        self.setWindowTitle(window_title)
        self.to_xml_button.setText(to_xml_button_text)
        self.to_mysql_button.setText(to_sql_button_text)
        self.excel_file_lbl.setText(excel_file_lbl)

    def button_on_load(self):
        # disable buttons
        self.read_excel_button.setDisabled(True)
        self.to_xml_button.setDisabled(True)
        self.to_mysql_button.setDisabled(True)
        self.copy_save_button.setDisabled(True)

        # read only text
        self.excel_data_output.setReadOnly(True)
        self.save_path_input.setReadOnly(True)
        self.xml_data_output.setReadOnly(True)
        self.sql_data_output.setReadOnly(True)

    def button_connection(self):
        # button function
        self.browse_button.clicked.connect(self.browse_files)
        self.read_excel_button.clicked.connect(self.read_excel)
        self.to_xml_button.clicked.connect(self.create_text_file)
        self.to_mysql_button.clicked.connect(self.to_mysql)
        self.copy_save_button.clicked.connect(self.copy_text)

    def browse_files(self):
        try:
            setting_last_file_path = QSettings("Order Helper", "LastSetting")  # path in regedit

            last_open_file_path = setting_last_file_path.value("folder_path")  # get value

            # open dialog box
            filepath = QFileDialog.getOpenFileName(self, 'Open File', last_open_file_path, "Excel File (*.xls *.xlsx)")

            if filepath[0] != "":
                file_path = path.normpath(filepath[0])

                self.file_path_input.setText(file_path)  # set value in line edit

                # split full path in folder name and file name from line edit
                folder_path, file_name = path.split(self.file_path_input.text())

                self.file_name_output.setText(str(file_name))

                setting_last_file_path.setValue("folder_path", folder_path)

                self.read_excel_button.setEnabled(True)
        except Exception as e:
            print(e)

    def read_excel(self):
        self.file_path_input.setDisabled(True)
        self.sheet_name_input.setDisabled(True)
        self.file_name_output.setDisabled(True)

        excel_file_path = self.file_path_input.text()
        sheet_name = self.sheet_name_input.text()

        # read excel sheet
        excel_data = read_excel_sheet(excel_file_path, sheet_name)

        if "error" not in excel_data:
            excel_data_string: str = f'Active Cell: {excel_data.get("active_cell")}\n' \
                                     f'Start Row: {excel_data.get("start_row")}\n' \
                                     f'Last Row: {excel_data.get("last_row")}\n' \
                                     f'Last Column: {excel_data.get("max_col")}\n' \
                                     f'WorkSheet: {excel_data.get("ws")}'

            self.ws = excel_data.get('ws')
            self.data_start_row: int = excel_data.get('start_row')
            self.data_last_row: int = excel_data.get('last_row')
            self.start_row_input.setText(str(self.data_start_row))
            self.last_row_input.setText(str(self.data_last_row))

            # activate button
            self.to_xml_button.setDisabled(False)
            self.to_mysql_button.setDisabled(False)
        else:
            excel_data_string: str = f"Error: {excel_data.get('error')}"

        self.excel_data_output.setText(str(excel_data_string))

    def create_text_file(self):
        try:
            self.start_row = int(self.start_row_input.text())
            self.last_row = int(self.last_row_input.text())

            # check user input validation
            input_validation = validate_user_data(self.start_row, self.last_row, self.data_start_row, self.data_last_row)
            if 'error' not in input_validation:
                # create xml string
                xml_data = None
                if self.sending_button == self.order_button:
                    xml_data = get_order_xml_file(self.ws, self.start_row, self.last_row)
                elif self.sending_button == self.return_button:
                    xml_data = get_return_xml_file(self.ws, self.start_row, self.last_row)

                # write text file
                xml_string = xml_data.get('xml_string')
                output_file_name = xml_data.get("file_name")
                folder_path = path.split(self.file_path_input.text())[0]

                text_location = create_text_file(xml_string, folder_path, output_file_name)
                self.save_path_input.setText(str(text_location))
                self.copy_save_button.setDisabled(False)
                xml_msg: str = f'{xml_data.get("number")} order xml generated\n' \
                               f'File Location: {text_location}'
            else:
                xml_msg: str = f'Error: {input_validation.get("error")}'

            self.xml_data_output.setText(str(xml_msg))
        except Exception as e:
            print(e)

    def copy_text(self):
        copy(self.save_path_input.text())
        self.copy_save_button.setText("Copied")

    def to_mysql(self):
        self.start_row = int(self.start_row_input.text())
        self.last_row = int(self.last_row_input.text())

        sql_validation = validate_sql_connection()
        input_validation = validate_user_data(self.start_row, self.last_row, self.data_start_row, self.data_last_row)
        if "error" not in sql_validation:  # check mysql connection
            mysqldb = sql_validation.get('mysqldb')  # get sql database
            if "error" not in input_validation:  # check for start row and last row
                sql_data = None
                if self.sending_button == self.order_button:  # check weather to pass order query or return query
                    sql_data = send_order_to_mysql(self.ws, self.start_row, self.last_row, mysqldb)
                elif self.sending_button == self.return_button:
                    sql_data = send_return_to_mysql(self.ws, self.start_row, self.last_row, mysqldb)

                if 'error' in sql_data:
                    sql_msg = sql_data.get('error')
                else:
                    sql_msg = f"{sql_data.get('number')} of order saved to database"
            else:
                sql_msg = input_validation.get('error')
        else:
            sql_msg = sql_validation.get('error')

        self.sql_data_output.setText(str(sql_msg))


##########################
##### MAIN FUNCTIONS #####
##########################

# check user input start and last row and returns error
def validate_user_data(start_row: int, last_row: int, data_start_row: int, data_last_row: int):
    # ws = excel worksheet,
    # start_row and last_row = user define data,
    # data_start_row and data_last_row = excel data
    if not isinstance(start_row, int) or not isinstance(last_row, int):
        return {'error': 'Invalid Input'}
    elif start_row < data_start_row:
        return {'error': "Default start row is 2"}
    elif start_row > data_last_row:
        return {'error': "start row cannot be more than excel last row"}
    elif last_row > data_last_row:
        return {'error': f"Last row in excel is {data_last_row}"}
    elif last_row < data_start_row:
        return {'error': "last row cannot be less than excel start row"}
    elif start_row > last_row:
        return {'error': 'Start row cannot be more than last row'}
    else:
        return {'msg': 'Valid Input'}


# create order xml string and returns xml string, number of data and output file name
def get_order_xml_file(ws, start_row: int, last_row: int):
    data_number = 0
    tally_company_id = ws.cell(row=start_row, column=18).value
    xml_string = head_xml(tally_company_id)
    for row in ws.iter_rows(min_row=start_row, max_row=last_row, max_col=ws.max_column, values_only=True):
        tally_vch_number = row[0]
        date = row[1]
        order_id = row[2]
        customer_name = row[3]
        gst_states_id = row[4]
        sku_id_with_color = row[5]
        sku_id = row[6]
        quantity = row[7]
        rate = row[8]
        shipping = row[9]
        cgst = row[10]
        sgst = row[11]
        igst = row[12]
        round_off = row[13]
        total = row[14]
        portal_name_id = row[15]
        # warehouse_id = row[16]
        # tally_company_id = row[17]

        date_format = datetime.strftime(date, '%Y%m%d')
        current_date_format = datetime.strftime(datetime.now(), '%d-%b-%Y at %H:%M')

        xml_string += order_body_xml(tally_vch_number, date_format, order_id, customer_name, gst_states_id,
                                     sku_id_with_color, sku_id, quantity, rate, shipping, cgst, sgst, igst, round_off,
                                     total, portal_name_id, current_date_format)
        data_number += 1
        print(f'{data_number} done')

    xml_string += tail_xml(tally_company_id)

    return {'xml_string': xml_string,
            'number': data_number,
            'file_name': "orderXML"}


# create return xml string and returns xml string, number of data and output file name
def get_return_xml_file(ws, start_row: int, last_row: int):
    data_number = 0
    tally_company_id = ws.cell(row=start_row, column=19).value
    xml_string = head_xml(tally_company_id)
    for row in ws.iter_rows(min_row=start_row, max_row=last_row, max_col=ws.max_column, values_only=True):
        return_date = row[0]
        tally_vch_number = row[1]
        return_order_id = row[2]
        return_type = row[3]
        design_name = row[4]
        design_number = row[5]
        piece = row[6]
        rate = row[7]
        shipping = row[8]
        cgst = row[9]
        sgst = row[10]
        igst = row[11]
        round_off = row[12] * -1
        total = row[13]
        order_date = row[14]
        portal_name = row[15]
        order_gst_state = row[16]
        customer_name = row[17]
        tally_company = row[18]
        order_data_id = row[19]
        return_initiated_id = row[20]
        penalty = row[21]

        return_date_format = datetime.strftime(return_date, '%Y%m%d')
        order_date_format = datetime.strftime(order_date, '%Y%m%d')
        current_date_format = datetime.strftime(
            datetime.now(), '%d-%b-%Y at %H:%M')

        xml_string += return_body_xml(return_date_format, tally_vch_number, return_order_id, return_type, design_name,
                                      design_number, piece, rate, shipping, cgst, sgst, igst, round_off, total,
                                      order_date_format, portal_name, order_gst_state, customer_name,
                                      current_date_format)

        data_number += 1
        print(f'{data_number} done')

    xml_string += tail_xml(tally_company_id)

    return {'xml_string': xml_string,
            'number': data_number,
            'file_name': "returnXML"}


# validate sql connection and return mysqldb or error
def validate_sql_connection():
    try:
        mysqldb = connect(
            host='192.168.0.2',  # '192.168.0.2'
            port='3306',  # 3306
            username='sunfashion',
            password='8632',
            database='online_database'
        )
        print(mysqldb)
    except Exception as e:
        return {'error': e}
    else:
        return {"mysqldb": mysqldb}


# send order to my sql and return number of data insert
def send_order_to_mysql(ws, start_row, last_row, mysqldb):
    cursor = mysqldb.cursor()

    # insert query
    def insert_order_data(tally_vch_number, date, order_id, customer_name, gst_states_id, sku_id_with_color, quantity,
                          rate, shipping, cgst, sgst, igst, round_off, total, portal_name_id, warehouse_id,
                          tally_company_id):
        args = [tally_vch_number, date, order_id, customer_name, gst_states_id, sku_id_with_color, quantity, rate,
                shipping, cgst, sgst, igst, round_off, total, portal_name_id, warehouse_id, tally_company_id]
        cursor.callproc('insert_order_data', args)

    # update query
    def update_order_data_rate(rate, shipping, cgst, sgst, igst, round_off, total, tally_vch_number):
        args = [rate, shipping, cgst, sgst, igst,
                round_off, total, tally_vch_number]
        cursor.callproc('update_order_data_rate', args)

    # main iter
    new_data_count = 0
    for row in ws.iter_rows(min_row=start_row, max_row=last_row, max_col=ws.max_column, values_only=True):
        tally_vch_number = str(row[0])
        date = row[1]
        order_id = str(row[2])
        customer_name = str(row[3])[:20]
        gst_states_id = str(row[4])
        sku_id_with_color = str(row[5])
        sku_id = str(row[6])
        quantity = row[7]
        rate = row[8]
        shipping = row[9]
        cgst = row[10]
        sgst = row[11]
        igst = row[12]
        round_off = row[13]
        total = row[14]
        portal_name_id = str(row[15])
        warehouse_id = str(row[16])
        tally_company_id = str(row[17])

        # insert query
        insert_order_data(tally_vch_number, date, order_id, customer_name, gst_states_id, sku_id_with_color, quantity,
                          rate, shipping, cgst, sgst, igst, round_off, total, portal_name_id, warehouse_id,
                          tally_company_id)

        # #update query
        # update_order_data_rate(rate, shipping, cgst, sgst,
        #                        igst, round_off, total, tally_vch_number)

        # commit
        mysqldb.commit()

        new_data_count += cursor.rowcount
        print(new_data_count)

    return {'number': new_data_count}


# send order to my sql and return number of data insert
def send_return_to_mysql(ws, start_row, last_row, mysqldb):
    cursor = mysqldb.cursor()
    new_data_count = 0
    for row in ws.iter_rows(min_row=start_row, max_row=last_row, max_col=ws.max_column, values_only=True):
        return_date = row[0]
        tally_vch_number = row[1]
        order_data_id = row[19]
        return_type = row[3]
        return_initiated_id = row[20]
        penalty = row[21]

        args = [return_date, tally_vch_number, order_data_id, return_type]
        cursor.callproc('insert_return_data', args)
        mysqldb.commit()

        new_data_count += cursor.rowcount
        print(new_data_count)

    return {'number': new_data_count}


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = OrderEntryWindow()

    sys.exit(app.exec_())

# entry_range = input(f'''All Data Range: {start_row}-{max_row}
# Active Row: {active_cell}
# Last Row: {max_row}
# Enter Data Range: ''')

# folder_path_ = "C:/Users/sunfashionLap/Desktop"
# file_name_ = "OrderXML2.xlsx"
# sheet_name_ = "For Tally"
#
# data=read_excel_sheet(folder_path_,file_name_,sheet_name_)
# print(data)
#
# xml = get_xml_file(data.get('ws'),data.get('start_row'),data.get('last_row'))
# create_text_file(xml.get('xml_string'),folder_path_,"test file")
