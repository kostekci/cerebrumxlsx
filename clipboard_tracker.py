#from tkinter import Tk  # Python 3
from tkinter import Tk

last_text=''

while True:
    text = Tk().clipboard_get()
    if text != last_text:
        last_text=text
        print(last_text)