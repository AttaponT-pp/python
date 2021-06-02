import os

from win32ctypes.pywin32.pywintypes import error
import win32con
import win32api

from time import sleep
from random import randint

from tkinter import Tk, Label, HORIZONTAL
from tkinter import ttk

from qq import img
import base64


def set_icon():
    ico_temp = "tmp.ico"
    tmp = open(ico_temp, "wb+")
    tmp.write(base64.b64decode(img))  # Write to temporary file
    tmp.close()
    root.iconbitmap(ico_temp)  # Set icon
    os.remove(ico_temp)  # Remove dying icon


def step_up(position):
    progress['value'] += 0.5
    sleep(0.1)
    try:
        position_next = win32api.GetCursorPos()
        if position != position_next:
            # clear progress bar
            progress['value'] = 0
    except error:
        pass


def clear_step():
    # clear progress bar
    progress['value'] = 0
    # get screen size
    x_max = win32api.GetSystemMetrics(0) - 1
    y_max = win32api.GetSystemMetrics(1) - 40
    pos_next = (randint(0, x_max), randint(0, y_max))
    # print('X = {}   Y = {}'.format(pos_next[0], pos_next[1]))
    # set new mouse position
    try:
        win32api.SetCursorPos(pos_next)
        # move mouse event
        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, pos_next[0], pos_next[1], 0, 0)
        # perform left click
        # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, pos_next[0], pos_next[1], 0, 0)
        sleep(0.1)
        win32api.SetCursorPos(pos_next)
        # perform right click
        # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, pos_next[0], pos_next[1], 0, 0)
        lbl1.config(text='X = {}   Y = {}'.format(pos_next[0], pos_next[1]))
    except error:
        print('Cannot set new mouse position.')


def mouse_move():
    # get mouse position
    try:
        pos = win32api.GetCursorPos()
        lbl1.config(text='X = {}   Y = {}'.format(pos[0], pos[1]))
        step_up(pos)
    except error:
        pass
    if progress['value'] > 100:
        clear_step()
    root.after(10, mouse_move)


# create Tkinter
root = Tk()
# change title name
root.title('MOUSE-MOVE')
# set application's icon
# root.iconbitmap('E:\\PythonProject\\MouseMove\\mouse.ico')
set_icon()
# set geometry and fix
root.geometry("250x45")
root.resizable(0, 0)
# add label
lbl1 = Label(root, text='')
lbl1.pack()
# add progress bar
progress = ttk.Progressbar(root, orient=HORIZONTAL, length=200, mode='determinate')
progress.pack()
# call function
mouse_move()
# Set application always on top
# root.wm_attributes("-topmost", 1)
# Hide the apps after launch
root.withdraw()
# update application gui
root.update()
# Reveal the application
root.iconify()
# launch the app
root.mainloop()
