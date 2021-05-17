import os
import ctypes
from pygetwindow import getAllTitles, getWindowsWithTitle
from time import sleep
from getpass import getuser

import numpy as np
import pywintypes
from win32com.client import GetObject

from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def dbschenker_login(URL, usr, pwd):
    # Navigate to Login Page
    browser.get(URL)
    minimize_console()
    print(browser.current_url)
    # minimize_console()
    try:
        # Find and click Login
        wait.until(EC.element_to_be_clickable((By.XPATH, '//es-nav-link[@eslogin]'))).click()
        sleep(5)
        # Fill Username
        wait.until(EC.element_to_be_clickable((By.ID, 'username'))).click()
        browser.find_element_by_id('username').send_keys(usr)
        # Fill Password
        wait.until(EC.presence_of_element_located((By.ID, 'userpassword'))).send_keys(pwd)
        # Click Sign in
        browser.find_element_by_xpath('//button/span[@id="signInButtonText"]').click()
        sleep(5)
        # wait until found InViewPro
        wait.until(EC.element_to_be_clickable((By.XPATH, '//div/a/i[@class="pvmi-inview-icon-64"]')))
        return True
    except TimeoutError:
        print('Time-out error')
        return False


def asn_upload(parent_hwnd, mpa_ids):
    print('Schenker Outbound ASN Upload.')
    print('Parent HWND: {}'.format(parent_hwnd))
    # Get all hwnd
    all_hwnd = browser.window_handles
    # Find InViewPro window
    for hwnd in all_hwnd:
        print('HWND : {}    Title name: {}'.format(hwnd, hwnd.title()))
        if hwnd != parent_hwnd:
            print('Found InViewPro window')
            browser.switch_to.window(hwnd)
            sleep(8)
            break
    # Find Transaction
    wait.until(EC.element_to_be_clickable((By.XPATH, '//li/a/span[contains(text(),"Transaction")]'))).click()
    # Select ASN/Receipt
    wait.until(EC.element_to_be_clickable((By.XPATH, '//li/div/a/span[contains(text(),"ASN/Receipt")]'))).click()
    return True
    # sleep(8)
    # Select MPa ID
    # for m in range(len(mpa_ids)):
    #     print(mpa_ids[m])
    #     wait.until(EC.element_to_be_clickable((By.XPATH, '//div/ng-select[@bindvalue="hubId"]'))).click()
    #     sleep(1)
    #     i = 0
    #     for hub_option in browser.find_elements_by_xpath('//div/div[@role="option"]'):
    #         print('Option: {}'.format(hub_option.text))
    #         if mpa_ids[i] in hub_option.text:
    #             hub_option.click()
    #             sleep(3)
    #             # check something that browser are completely loaded
    #             wait.until(EC.element_to_be_clickable((By.XPATH, '//div/button[@title="Upload"]'))).click()
    #             sleep(3)
    #     i += 1
    # # Find and click upload ASN file
    # wait.until(EC.presence_of_element_located((By.XPATH, '//div/button[@title="Upload"]'))).click()
    # # Wait until app-upload popup
    # wait.until(EC.presence_of_element_located((By.XPATH, '//div/div/span[@id="ui-dialog-5-label"]')))
    # # Browse ASN file
    # browser.find_element_by_xpath('//div/div[@class="uploadfilecontainer"]').click()
    # # Wait until file is fully uploaded.
    # try:
    #     wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="files-list ng-star-inserted"]/p[contains(text(), ".xlsx")]')))
    #     file_name = browser.find_element_by_xpath('//div[@class="files-list ng-star-inserted"]').text
    #     print('ASN File: {}'.format(file_name))
    #     # Click upload
    #     browser.find_element_by_xpath('//div/button[contains(text(), "Upload")]').click()
    #     sleep(3)
    #     return True
    # except TimeoutError:
    #     return False


def get_asn_template(path):
    vba_file = ''
    asn_file = ''
    wb = ''
    ws = ''
    user_data = ''
    mpa_id = ''
    print('Application path: {}'.format(path))
    files = os.listdir(path)
    for idx in range(len(files)):
        print(files[idx])
        if '~$' in files[idx]:
            pass
        elif 'Schenker ASN Upload' in files[idx] and '.xlsm' in files[idx]:
            vba_file = files[idx]
            print('Found VBA ASN Upload: {}'.format(files[idx]))
        elif 'ASN_CISCO_template' in files[idx] and '.xlsx' in files[idx]:
            asn_file = files[idx]
            print('Found ASN template file: {}'.format(files[idx]))
    if not vba_file:
        print('Not found file VBA file {}'.format(vba_file))
    elif not asn_file:
        print('Not found file ASN file {}'.format(asn_file))
    try:
        xl = GetObject(None, "Excel.Application")
        wb = xl.Workbooks(vba_file)
    except pywintypes.com_error:
        print('File: {} does not open.\nPlease open {}'.format(vba_file, vba_file))
        mbox(title='Scripts Error', text='File: {} does not open.\nPlease open {}'.format(vba_file, vba_file), style=0)
    for s in range(len(wb.Sheets)):
        print(wb.Sheets[s].Name)
        if 'HEADER' in wb.Sheets[s].Name:
            ws_header = wb.Sheets[s]
            mpa_id_list = []
            for i in range(2, len(ws_header.UsedRange.Rows)+1):
                print(ws_header.Cells(i, 13).Value)
                mpa_id_list.append(ws_header.Cells(i, 13).Value)
            mpa_id = np.unique(mpa_id_list)
        elif 'ASN list' in wb.Sheets[s].Name:
            print('Found {}'.format(wb.Sheets[s].Name))
            ws_asn = wb.Sheets[s]
            url = ws_asn.Cells(16, 11).Value
            user = ws_asn.Cells(17, 11).Value
            password = ws_asn.Cells(18, 11).Value
            user_data = [url, user, password]
    output_data = [user_data, mpa_id]
    return output_data


def validate_asn_upload():
    print('*** Validate ASN Upload. ***')
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, '//div/span[contains(text(), "Error Detail")]')))
        print('Error Detail')
        print(browser.find_element_by_xpath('//tr/th[@id="detailHeader"]'))
        for x in browser.find_elements_by_xpath('//tbody/tr[class@ng-star-inserted]/td'):
            print(x)
        sleep(2)
        browser.find_element_by_css_selector("*a.ng-tns-c4-32").click()
    except TimeoutError:
        print('Not Found the error.')
    return True


def mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def minimize_console():
    window = getAllTitles()
    for i in window:
        if 'geckodriver' in i:
            print('found!!')
            driver_console = getWindowsWithTitle(i)[0]
            driver_console.minimize()


def kill_console():
    window = getAllTitles()
    for i in window:
        if 'geckodriver' in i:
            print('found!!')
            driver_console = getWindowsWithTitle(i)[0]
            driver_console.close()


if __name__ == '__main__':
    # Setup save path
    # save_path = "C:\\Temp\\outbound_asn_upload"
    save_path = "E:\\PythonDev\\outbound_asn_upload"
    # application's path
    app_path = os.getcwd()
    if app_path != save_path:
        os.chdir(save_path)
        print(os.getcwd())
        app_path = os.getcwd()
    print('App_Path: {}'.format(app_path))
    # get data from VBA file
    input_data = get_asn_template(app_path)
    print('Data" {} TYPE: {}'.format(input_data, type(input_data)))
    print('Input data = {} \nMPA ID = {}'.format(input_data[0], input_data[1]))
    # Init Web Driver, set firefox capabilities
    print('*** Initial FireFox Web-Driver ***')
    firefox_capabilities = DesiredCapabilities.FIREFOX
    firefox_capabilities['marionette'] = True
    # Identify FireFox's profile path
    ffProfilePath = 'C:\\Users\\' + getuser() + '\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\bjyghxs8.Test'
    # Define firefox profile
    profile = webdriver.FirefoxProfile(profile_directory=ffProfilePath)
    geckoPath = 'C:\\Python39\\geckodriver\\geckodriver.exe'
    browser = webdriver.Firefox(firefox_profile=profile, capabilities=firefox_capabilities, executable_path=geckoPath)
    # Init Delay
    wait = WebDriverWait(browser, 120)
    # Login to E-DBSchenker
    if not dbschenker_login(input_data[0][0], input_data[0][1], input_data[0][2]):
        print('Unsuccessful E-Schenker login, Please contact engineer.')
    # Get parent's window handling
    parent_hwnd = browser.current_window_handle
    # Launch InViewPro
    browser.find_element_by_xpath('//div/a/i[@class="pvmi-inview-icon-64"]').click()
    sleep(10)
    print('Parent HWND: {}'.format(parent_hwnd))
    # Upload ASN
    if not asn_upload(parent_hwnd, input_data[1]):
        print('No ASN file selected.')
        kill_console()
    else:
        mbox(title="Outbound ASN Upload", text="Access to DBSchenker Completed.\nPlease select ASN file.", style=0)
        kill_console()
    # Validate ASN Upload Results
    # validate_asn_upload()