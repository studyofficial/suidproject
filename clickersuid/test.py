import pyautogui
from time import sleep
try:
    x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\date.PNG')
    pyautogui.doubleClick(x, y)
except:
        pass
sleep(1)
try:

    x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\copy.PNG')
    pyautogui.click(x, y)
except:
        pass