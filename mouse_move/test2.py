import ctypes
from time import sleep
from win32 import win32gui
from win32.lib import win32con


def winEnumHandler(hwnd, ctx):
    if win32gui.IsWindowVisible( hwnd ):
      print (hex(hwnd), win32gui.GetWindowText( hwnd ))
      callback(hwnd, None)
      # active_app = win32gui.GetWindowText(hwnd)
      # root_app_name = ['Auto Collection', 'UDTS', 'FITS-Data', 'MOUSE-MOVE', 'Outlook', 'Visual Studio Code']     
      # app_dicts = {}
      # for i in range(0, len(root_app_name)):      
      #   if root_app_name[i] in active_app:
      #     app_window = callback(hwnd, None)
      #     app_dicts[root_app_name[i]] = app_window
      #     print(app_dicts)
        #   print('Position X = {},\nPosition Y = {}'.format(app_window['app_position'][0], app_window['app_position'][1]))
        #   print('App Width = {},\nApp Height = {}'.format(app_window['app_dimension'][0], app_window['app_dimension'][1]))
          # if win32gui.IsWindowVisible(hwnd):
            # win32gui.ShowWindow(hwnd, 5)
            
            # win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            # win32gui.UpdateWindow(hwnd)
            # win32gui.BringWindowToTop(hwnd)
            # win32gui.SetForegroundWindow(hwnd)
          #   sleep(1)
            
def callback(hwnd, extra):
    rect = win32gui.GetWindowRect(hwnd)
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y
    x_min = rect[0]
    y_min = rect[1]
    x_max = rect[2]
    y_max = rect[3]
    # return {"app_position": [x, y], "app_size": [x_min, x_max, y_min, y_max]}
    # print("Window %s:" % win32gui.GetWindowText(hwnd))
    print("\tLocation: (%d, %d)" % (x, y))
    print("\t    Size: (%d, %d)" % (w, h))
    return {"app_position": [x, y], "app_dimension": [w, h], "app_size": [x_min, x_max, y_min, y_max]}
  
win32gui.EnumWindows( winEnumHandler, None )
