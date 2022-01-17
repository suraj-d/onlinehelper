import keyboard
import time
from tkinter import Tk

data = Tk().clipboard_get()
data = data.split("\n")

time.sleep(3)


def paste_data(data: list, repeat: int = 1):

    if len(data) > 1:
        data.pop(-1)

    for r in range(repeat):
        for d in data:
            keyboard.write(d)
            time.sleep(0.5)
            keyboard.press_and_release('enter')


# add number of time to repeat
# paste_data(date, 2)
paste_data(data, 160)
