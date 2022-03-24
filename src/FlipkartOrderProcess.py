from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import PyPDF2
import os


options = webdriver.ChromeOptions()
downloadPath = r"C:\test\picklist\batch"
prefs = {}
os.makedirs(downloadPath, exist_ok=True)
prefs["profile.default_content_settings.popups"] = 0
prefs["download.default_directory"] = downloadPath
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(
    service=Service(r"D:\PycharmProjects\OrderToMySQL\chromedriver.exe"), options=options)

driver.get("https://seller.flipkart.com/index.html#dashboard/fbflite-ff")

# wait till username visible
WebDriverWait(driver, 20).until(
    EC.visibility_of_element_located((By.NAME, "username"))).send_keys(
    'sunfashionandlifestyle@gmail.com')
driver.find_element(By.NAME, 'username').send_keys(Keys.RETURN)

# wait till password visible
WebDriverWait(driver, 20).until(
    EC.visibility_of_element_located((By.NAME, "password"))).send_keys('suraj@123')
driver.find_element(By.NAME, 'password').send_keys(Keys.RETURN)


def get_picklist():
    picklist_data = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'picklist-table-data'))).text

    picklist_data_list = picklist_data.split('\n')
    pick_list = []
    number = 0
    for i in range(0, len(picklist_data_list), 9):
        select_picklist = {}
        number += 1
        select_picklist['number'] = number
        select_picklist['picklist_id'] = picklist_data_list[i].strip()
        select_picklist['type'] = picklist_data_list[i + 1].strip()
        select_picklist['date'] = picklist_data_list[i + 2].strip()
        select_picklist['time'] = picklist_data_list[i + 3].strip()
        select_picklist['status'] = picklist_data_list[i + 4].strip()
        select_picklist['order'] = picklist_data_list[i + 5].strip()
        select_picklist['pending_order'] = picklist_data_list[i + 6].strip()
        select_picklist['label_ready'] = picklist_data_list[i + 7].strip()

        pick_list.append(select_picklist)
    return pick_list


# pick_list = [{'number': 1, 'picklist_id': 'P251121-9BFBC38EACEE', 'type': 'REGULAR', 'date': 'Nov 25, 2021', 'time': '07:06', 'status': 'Open', 'order': '4', 'pending_order': '4', 'label_ready': '0'},
#              {'number': 2, 'picklist_id': 'P251121-BD44C276EDE6', 'type': 'REGULAR', 'date': 'Nov 25, 2021', 'time': '07:41', 'status': 'Open', 'order': '2', 'pending_order': '2', 'label_ready': '0'}]


# # get selected picklist from all this generated picklist
# def select_picklist_from_input(picklist: list):
#     for i in picklist:  # print number and picklist for input details
#         number = i.get('number')
#         pick_list_id = i.get('picklist_id')
#         print(f'{number}: {pick_list_id}')
#
#     pick_list_input_number = str(
#         input('Enter number to download picklist (1,3): '))
#
#     if "," in pick_list_input_number:  # check if input contant multiple input value
#         pick_list_input_number = pick_list_input_number.split(',')
#     else:
#         pick_list_input_number = pick_list_input_number
#
#     selected_picklist = []
#     for i in picklist:  # loop throught dict detail in pick_list
#         number = i.get('number')
#         pick_list_id = i.get('picklist_id')
#         for j in pick_list_input_number:  # loop through pick_list_input_number list
#             if int(j) == number:
#                 selected_picklist.append(pick_list_id)
#
#     return selected_picklist


# get the list of all the lsn from pdf_path
def get_lsn_from_pdf(pdf_path: str):
    """

    :param pdf_path:
    :return: list of lsn, sku and order qty
    """
    # pdf_path = r"C:\Users\sunfashionLap\Desktop\Flipkart-PickList-P020921-4A79C9F22CB8-02-Sep-2021-06-15.pdf"
    doc = fitz.open(pdf_path)

    # get text from each pdf page
    for page in doc:
        text = page.get_text()

    text_list = text.split("\n")
    order_list = []

    for i in range(len(text_list)):
        list_dict = {}
        single_text = text_list[i].strip()
        if 'lst' in single_text.lower():
            lsn = single_text

        if 'sku' in single_text.lower():
            list_dict['lsn'] = lsn
            list_dict['sku'] = single_text
            list_dict['orders'] = int(text_list[i + 1])
            order_list.append(list_dict)

    full_list_of_lsn = []
    for detail in order_list:
        for x in range(detail.get('orders')):
            full_list_of_lsn.append(detail.get('lsn'))

    # print(full_list_of_lsn)
    return full_list_of_lsn


picklist_path = 'S:\Flipkart-PickList-P010122-46A7A8D5A32E-01-Jan-2022-10-00.pdf'
a = get_lsn_from_pdf(picklist_path)
print(a)

'''TODO: 
1. click pick list
2. paste lsn in box
3. download label
4. get tracking from label
5. loop through lsn list in dispatch'''

# #click on pick list id for dispatch

# WebDriverWait(driver, 20).until(
#     EC.visibility_of_element_located((By.CSS_SELECTOR, f"[data-picklist-id = '{i}']"))).click()

# driver.close()
