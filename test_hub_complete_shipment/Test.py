# import os
# import psutil
# import win32gui
# from FITS_Connect import *
#
# # Ref.: https://gist.github.com/Sanix-Darker/8cbed2ff6f8eb108ce2c8c51acd2aa5a
# def checkIfProcessRunning(processName):
#     '''
#     Check if there is any running process that contains the given name processName.
#     '''
#     # Iterate over the all the running process
#     for proc in psutil.process_iter():
#         try:
#             # Check if process name contains the given name string.
#             if processName.lower() in proc.name().lower():
#                 return True
#         except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#             pass
#     return False;
#
# # Check if any chrome process was running or not.
# # appname = "ShipmentRTV"
# # if checkIfProcessRunning(appname):
# #     print('Yes a ' + appname + ' process was running')
# # else:
# #     print('No, '+ appname + ' process was running')
# #     os.startfile("C:\Program Files\Cisco\ShipmentRTV\ShipmentRTV.exe")
#
# def dumpWindows(hwnd, _windowEnumerationHandler=None):
#     """Dump all controls from a window
#
#     Useful during development, allowing to you discover the structure of the
#     contents of a window, showing the text and class of all contained controls.
#
#     Parameters
#     ----------
#     hwnd
#         The window handle of the top level window to dump.
#
#     Returns
#     -------
#         all windows
#
#     Usage example::
#
#         replaceDialog = findTopWindow(wantedText='Replace')
#         pprint.pprint(dumpWindow(replaceDialog))
#     """
#     windows = []
#     win32gui.EnumChildWindows(hwnd, _windowEnumerationHandler, windows)
#     return windows
#
#
#
# MAIN_HWND = 0
#
# def is_win_ok(hwnd, starttext):
#     s = win32gui.GetWindowText(hwnd)
#     if s.startswith(starttext):
#             print(s)
#             global MAIN_HWND
#             MAIN_HWND = hwnd
#             return None
#     return 1
#
#
# def find_main_window(starttxt):
#     global MAIN_HWND
#     win32gui.EnumChildWindows(0, is_win_ok, starttxt)
#     return MAIN_HWND
#
#
# def winfun(hwnd, lparam):
#     s = win32gui.GetWindowText(hwnd)
#     if len(s) > 3:
#         print("winfun, child_hwnd: %d   txt: %s" % (hwnd, s))
#     return 1
#
#
# def main():
#     main_app = 'RTV Shipment'
#     # hwnd = win32gui.FindWindow(None, main_app)
#     hwnd = win32gui.FindWindow(main_app, None)
#     print(hwnd)
#     if hwnd < 1:
#         y = 1
#         print(y)
#         hwnd = find_main_window(main_app)
#     print(hwnd)
#     if hwnd:
#         y = 2
#         print(y)
#         win32gui.EnumChildWindows(hwnd, winfun, None)
# main()
#
# from openpyxl import *
# import re
# from PyQt5 import QtWidgets

# dlg = QtWidgets.QFileDialog.getOpenFileName()
# f_path = 'E:\\PythonDev\\test_hub_complete_shipment\\OBA_Request on 05-Mar-2021.xlsx'
# f_path = 'E:/PythonDev/test_hub_complete_shipment/OBA_Request on 05-Mar-2021.xlsx'
# f_path  = 'E:/PythonDev/test_hub_complete_shipment/Label RTV 10 Mar.xls'

# if f_path.endswith('.xlsx'):
#     print('Found .xlsx file')
# elif f_path.endswith('.xls'):
#     print('Found .xls file')
# else:
#     print('Not excel file')

# if re.search('.xlsx', f_path, re.IGNORECASE):
#     print('Found .xlsx file')
# elif re.search('.xls', f_path, re.IGNORECASE):
#     print('Found .xls file')
# else:
#     print('Not excel file')

# wb = load_workbook(f_path)
# ws = wb.active
# print(ws.title)
# print(ws.max_row)
# input_inv = '2112132'
# input_etr = 'R20210130'
# for row in range(5, ws.max_row + 1):
#     if ws.cell(row=row, column=10).value is None:
#         break
#     inv_num = str(ws.cell(row=row, column=10).value)
#     inv_num = str(ws.cell(row=row, column=10).value)
#     etr_num = ws.cell(row=row, column=23).value
#     print(row)
#     print(inv_num)
#     print(inv_num.split('\n')[0])
#     if re.search(input_inv, inv_num, re.IGNORECASE):
#         print('Found Invoice# {}'.format(inv_num.split('\n')[0]))
#         break
# import sys		
# from PyQt5.QtWidgets import QApplication, QPushButton
# from PyQt5.QtCore import QTimer


# class AppDemo(QPushButton):
# 	def __init__(self):
# 		super().__init__('My Button')
# 		self.resize(400, 400)
# 		self.setStyleSheet('font-size: 40px;')
# 		self.flag = True

# 		timer = QTimer(self, interval=1000)
# 		timer.timeout.connect(self.flashing)
# 		timer.start()

# 		self.clicked.connect(lambda: print('hello world'))

# 	def flashing(self):
# 		if self.flag:
# 			self.setStyleSheet('background-color: none; font-size: 40px')
# 		else:
# 			self.setStyleSheet('background-color: orange; font-size: 40px')

# 		self.flag = not self.flag


# if __name__ == '__main__':
# 	app = QApplication(sys.argv)

# 	demo = AppDemo()
# 	demo.show()
	
# 	try:
# 		sys.exit(app.exec_())
# 	except SystemExit:
# 		print('Closing Window...')


# from FITS_Connect import *

# model = '*'
# rev = '2.9.0.0'
# fits_dll = client.Dispatch("FITSDLL.clsDB")

# init = fits_dll.fn_InitDB('*', model, rev)
# if not init:
#     print('Init FITSDLL Error.')
#     exit()

# hand_check = fits_dll.fn_handshake(model, '1501', rev, '026487')
# status = fits_dll.fn_query(model, '1501', rev, 'EN', '026487')
# print('Hand check: {}'.format(hand_check))
# print('Query status: {}'.format(status))
which_opn = {'1501': True, '702': True, '1502': True, '1801': True}
for i in which_opn.items():
    print(i)
    if  i[1] == True:
        print(i[0])
        