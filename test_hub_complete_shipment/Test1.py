# import ctypes
# from time import sleep
# from win32 import win32gui
# from win32.lib import win32con

# def winEnumHandler(hwnd, ctx):
#     if win32gui.IsWindowVisible( hwnd ):
#       print (hex(hwnd), win32gui.GetWindowText( hwnd ))
#       callback(hwnd, None)
#       active_app = win32gui.GetWindowText(hwnd)
#       root_app_name = ['Shipping for Cisco Sustaining']     
#       app_dicts = {}
#       for i in range(0, len(root_app_name)):      
#         if root_app_name[i] in active_app:
#           app_window = callback(hwnd, None)
#           app_dicts[root_app_name[i]] = app_window
#         #   print(app_dicts)
#           print('Position X = {},\nPosition Y = {}'.format(app_window['app_position'][0], app_window['app_position'][1]))
#         #   print('App Width = {},\nApp Height = {}'.format(app_window['app_dimension'][0], app_window['app_dimension'][1]))
#           if win32gui.IsWindowVisible(hwnd):
#             win32gui.ShowWindow(hwnd, 5)
#             win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
#             win32gui.UpdateWindow(hwnd)
#             win32gui.BringWindowToTop(hwnd)
#             win32gui.SetForegroundWindow(hwnd)
#             sleep(1)
        
            
# def callback(hwnd, extra):
#     rect = win32gui.GetWindowRect(hwnd)
#     x = rect[0]
#     y = rect[1]
#     w = rect[2] - x
#     h = rect[3] - y
#     x_min = rect[0]
#     y_min = rect[1]
#     x_max = rect[2]
#     y_max = rect[3]
#     # return {"app_position": [x, y], "app_size": [x_min, x_max, y_min, y_max]}
#     # print("Window %s:" % win32gui.GetWindowText(hwnd))
#     # print("\tLocation: (%d, %d)" % (x, y))
#     # print("\t    Size: (%d, %d)" % (w, h))
#     return {"app_position": [x, y], "app_dimension": [w, h], "app_size": [x_min, x_max, y_min, y_max]}
  
# win32gui.EnumWindows( winEnumHandler, None )

# 
# from time import sleep
# from win32 import win32gui
# from win32 import win32api
# from win32.lib import win32con
# import pyautogui

# MAIN_HWND = 0

# def is_win_ok(hwnd, starttext):
#     s = win32gui.GetWindowText(hwnd)
#     if s.startswith(starttext):
#             print(s)
#             global MAIN_HWND
#             MAIN_HWND = hwnd
#             return None
#     return 1


# def find_main_window(starttxt):
#     global MAIN_HWND
#     win32gui.EnumChildWindows(0, is_win_ok, starttxt)
#     return MAIN_HWND


# def winfun(hwnd, lparam):
#     s = win32gui.GetWindowText(hwnd)
#     if len(s) > 3:
#         windowRec = win32gui.GetWindowRect(hwnd)
#         print("winfun, child_hwnd: %d   txt: %s" % (hwnd, s))
#         print(windowRec)
#         if 'Start' in s:
#             print('Start')
#             win32api.SendMessage(hwnd, win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
#             sleep(1)
#             win32api.SendMessage(hwnd, win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
#     return 1

# def main():
#     # main_app = 'Shipping for Cisco Sustaining'
#     main_app = 'Notepad++'
#     # hwnd = win32gui.FindWindow(None, main_app)
#     hwnd = win32gui.FindWindow(main_app, None)
#     print(hwnd)
#     if hwnd < 1:
#         hwnd = find_main_window(main_app)
#     print(hwnd)
#     if hwnd:
#         win32gui.EnumChildWindows(hwnd, winfun, None)

# main()
# import psutil
# from pywinauto.application import Application
# import pyautogui

# hwnds = pyautogui.getAllWindows()
# # print(hwnds)
# titles = pyautogui.getAllTitles()
# for i in range(0, len(titles)):
#     handle = pyautogui.getWindowsWithTitle(titles[i])
#     if titles[i] != "":
#         # print('Title: {} HWND: {}'.format(titles[i], handle))
#         if 'Shipping for Cisco Sustaining' in titles[i]:
#             print(handle)
#             print(type(handle))
            

# app = Application(backend="uia").start('C:\Program Files\Cisco\shipment2\Shipment.exe', timeout=10)
# # app = Application().connect(title='Untitled - Notepad')
# # describe the window inside Notepad.exe process
# dlg_spec = app.ShippingforCiscoSustaining
# # wait till the window is really open
# actionable_dlg = dlg_spec.wait('visible')

# dlg_spec.wrapper_object()
# dlg_spec.print_control_identifiers()
# dlg_spec.Start.Click()

import os
from win32 import win32api, win32gui
from pywinauto.application import Application
import pyautogui as gui


titles = gui.getAllTitles()
for i in range(0, len(titles)):
    print(titles[i])
    title = titles[i]
    if 'Shipping for Cisco Sustaining' in title:
        app = Application("uia").connect(title_re=".*Shipping for Cisco Sustaining*", timeout=10)
        break
    if i == len(titles) - 1:
        app = Application(backend="uia").start("C:\Program Files\Cisco\shipment2\Shipment.exe",timeout=10)
        break

main_window = app.window(title=title)
# main_window.child_window(title='Start')
# main_window.Start.click()
# main_window.print_control_identifiers()