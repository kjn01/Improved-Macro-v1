import tkinter as tk
from tkinter import *
from tkinter import simpledialog
import keyboard
import win32gui
# import win32api
from pywinauto import application
import time
from pywinauto import findwindows
import regex as re
import threading

ROOT = tk.Tk()

# ROOT.withdraw()

# inp = simpledialog.askstring(title="input test", prompt="type something: ") # opens new pop-up

# def setInf(val):
#     intN = None
#     if val == "1000":
#         intN = Label(ROOT, text="Infinite")
#         intN.pack()
#     else:
#         intN = Label(ROOT, text=str(val))
#         intN.pack()

title = None

def winEnumHandler(hwnd, ctx):
    global title
    if win32gui.IsWindowVisible(hwnd):
        if "Notepad" in win32gui.GetWindowText(hwnd):
            title = re.escape(win32gui.GetWindowText(hwnd))
            print(title)

def startMacro(txt, interval, loops):
    global title
    global start
    if start["text"] == "Start":
        start["text"] = "Stop"
    else:
        
        return
    win32gui.EnumWindows(winEnumHandler, None)
    app = application.Application(backend="win32").connect(best_match=title)
    form = app.window(title_re=title)
    # print(findwindows.find_windows(form))
    # form.print_control_identifiers()
    # form.Edit.send_keystrokes("hiei")
    for i in range(loops):
        form.Edit.send_keystrokes(txt)
        time.sleep(interval)


def macroThread(txt, interval, loops):
    threading.Thread(target=lambda: startMacro(txt, interval, loops)).start()

msgLabel = Label(ROOT, text="Enter message below")
msgLabel.pack()
msg = Text(ROOT, height=3, width=20)
msg.pack()

intLabel = Label(ROOT, text="Enter interval")
intLabel.pack()
interval = Scale(ROOT, from_=0, to=1000, orient=HORIZONTAL, length=500)
interval.pack()

durLabel = Label(ROOT, text="Enter duration")
durLabel.pack()
duration = Scale(ROOT, from_=1, to=1000, orient=HORIZONTAL, length=500)
duration.pack()

start = Button(ROOT, text="Start", command=lambda: macroThread(msg.get("1.0",'end-1c'), int(interval.get()), int(duration.get())))
start.pack()

mainloop()
