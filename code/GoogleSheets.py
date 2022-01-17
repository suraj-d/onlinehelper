from webbrowser import open
from os import system


def open_costing_sheet():
    open("https://docs.google.com/spreadsheets/d/1CwsnXJ4Qj-FNVtJOnVt2huO4GnHYLEf0/edit#gid=263504063")


def open_price_calculator():
    open("https://docs.google.com/spreadsheets/d/1CwsnXJ4Qj-FNVtJOnVt2huO4GnHYLEf0/edit#gid=1579262232")


def open_stock_sheet():
    open("https://docs.google.com/spreadsheets/d/10p92AlDT4ifJErbrKi-HoGliiWZlMN_DwrOVz0YCTCk/edit#gid=0")


def shutdown_server():
    system(r"shutdown /s /m \\sunserver")


def restart_server():
    system(r'shutdown /r /m \\sunserver')


