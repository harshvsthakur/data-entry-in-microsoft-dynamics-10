import pyautogui, glob, os, keyboard, time, pygetwindow as gw

x = input ("Enter DO Number : ")
y = input ("Enter DO Date in DDMMYY : ")
num = input ("Enter 5 digit PO Number to Open : ")
while len(num) != 5 or (not num.isdigit()):
    print ("Not a 5 digit number !")
    num = input("Please enter your 5 digit number : ")

num = int(num)

if ('Microsoft Dynamics GP - \\\\Remote') in gw.getAllTitles():
    time.sleep(0.1)
else:
    os.system("C:\\Users\\harshvardhans\\Desktop\\python\\1.py")

while True:
    if ('Microsoft Dynamics GP - \\\\Remote') in gw.getAllTitles():
        gpWindow = gw.getWindowsWithTitle('Microsoft Dynamics GP - \\\\Remote')[0]
        gpWindow.restore()
        gpWindow.activate()
        break
    else: time.sleep (1)

time.sleep(1)

keyboard.press_and_release('alt+a,p,r,enter')  # opens receivings transactions

while True:
    if ('.Receivings Transaction Entry. - \\\\Remote') in gw.getAllTitles():
        RecWindow = gw.getWindowsWithTitle('.Receivings Transaction Entry. - \\\\Remote')[0]
        RecWindow.restore()
        RecWindow.activate()
        break
    else: time.sleep (1)

time.sleep(1)

keyboard.press_and_release('tab,ctrl+c,tab')
pyautogui.typewrite(x)   # INPUT Invocie Number
keyboard.press_and_release('tab')
pyautogui.typewrite(y)   # INPUT Invoice Date
keyboard.press_and_release('tab')
pyautogui.typewrite("har")
keyboard.press_and_release('ctrl+v,tab,enter,enter,alt+u')
pyautogui.typewrite("MEP00"+str(num))   # INPUT PO Number
keyboard.press_and_release('tab,alt+v')

time.sleep(1)

keyboard.press_and_release('tab,tab,tab,tab')

os.chdir("Z:\\Purchase Dept 17112011\\Scanned LPOs\\ARJ PO - MEP")
for file in glob.glob(str(num)+"*.pdf"):
    filepath = os.getcwd()+"\\"+file
    os.startfile(filepath)
