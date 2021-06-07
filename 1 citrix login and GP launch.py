'''This Script opens the web browser, logs in to the GP and open PO entry 
window '''

import time, glob, os, pygetwindow as gw, pyautogui, keyboard

user = os.getlogin()
user_pc_pass = input('Enter computer password : ')
user_gp_pass = input('Enter GP password : ')

from webbot import Browser #Open web page,login, and download citrix client
web = Browser()
url = r'http://apps.arjgroup.com/Citrix/AccessPlatform/auth/login.aspx'
web.go_to(url) 
web.type(user , into='User name')
web.type(user_pc_pass , into='Password' ) 
web.click('Log In')
web.go_to('http://apps.arjgroup.com/Citrix/AccessPlatform/site/launcher.aspx?NFuse_Application=Citrix.MPS.App.CitrxNew.MEP%20GP&LaunchId=1569914789283')

time.sleep(5)

list_of_files = glob.glob('C:\\Users\\'+os.getlogin()+'\\Downloads\\*.ica')
latest_file = max(list_of_files, key=os.path.getctime)
os.startfile(latest_file)

def get_window(name):
    '''Gets window in front by gw.getAllTitles as argument'''
    while True:
        if (name) in gw.getAllTitles():
            gpWindow = gw.getWindowsWithTitle(name)[0]
            gpWindow.restore()
            gpWindow.activate()
            break
        else: time.sleep (1)


get_window ('Welcome to Microsoft Dynamics GP - \\\\Remote')

pyautogui.typewrite(os.getlogin())
keyboard.press_and_release('tab')
pyautogui.typewrite(user_gp_pass)
keyboard.press_and_release('enter')

get_window ('Company Login - \\\\Remote')
keyboard.press_and_release('down,down,enter')

time.sleep (2)

if gw.getAllTitles().count('Microsoft Dynamics GP - \\\\Remote') > 1:
    time.sleep(1)
    keyboard.press_and_release('alt+y,alt+d,enter')

time.sleep(15)

keyboard.press_and_release('alt+a,p,enter')
