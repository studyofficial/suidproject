from time import sleep

import pyautogui, sys

# Указание количества запросов
# requests = int(input("Количество запросов - "))
requests = 100
pyautogui.PAUSE = 0.6
pyautogui.FAILSAFE = True

# путь к файлам png
path = r'C:\Users\008an\PycharmProjects\suidproject\png/'
# номер запроса
requestnum = 0

pngname = "Name"


def find_img(pngname):
    sleep(0.5)
    r = None
    count = 50
    while r is None:
        count -= 1
        try:
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.9)
        except pyautogui.ImageNotFoundException:
            pass
        if count == 0:
            print(pngname + 'не найден')
            sys.exit()
            break
        else:
            if r != None:
                return r
            else:
                continue

def find_img_fast(pngname):
    sleep(0.5)
    r = None
    count = 10
    while r is None:
        count -= 1
        try:
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.9)
        except pyautogui.ImageNotFoundException:
            pass
        if count == 0:
            print(pngname + 'не найден')
            sys.exit()
            break
        else:
            if r != None:
                return r
            else:
                continue

def find_click(pngname):
    # Ожидаем появление элемента
    sleep(0.5)
    r = None
    count = 300
    while r is None:
        count -= 1
        try:
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.8)
        except pyautogui.ImageNotFoundException:
            pass
        if count == 0:
            print(pngname + 'не найден')
            sys.exit()
            break
        else:
            if r != None:
                pyautogui.click(r)
                return r
            else:
                continue


def find_click_kind(pngname):
    # Ожидаем появление элемента
    sleep(0.5)
    r = None
    count = 200
    while r is None:
        count -= 1
        try:
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.7)
        except pyautogui.ImageNotFoundException:
            pass
        if count == 0:
            print(pngname + 'не найден')
            sys.exit()
            break
        else:
            if r != None:
                pyautogui.click(r)
                return r
            else:
                continue

def find_click_exact(pngname):
    # Ожидаем появление элемента
    sleep(0.5)
    r = None
    count = 200
    while r is None:
        count -= 1
        try:
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.9)
        except pyautogui.ImageNotFoundException:
            pass
        if count == 0:
            print(pngname + 'не найден')
            break
        else:
            if r != None:
                pyautogui.click(r)
                return r
            else:
                continue

def find_click_exactly(pngname):
    # Ожидаем появление элемента
    sleep(0.5)
    r = None
    count = 200
    while r is None:
        count -= 1
        try:
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.99)
        except pyautogui.ImageNotFoundException:
            pass
        if count == 0:
            print(pngname + 'не найден')
            break
        else:
            if r != None:
                pyautogui.click(r)
                return r
            else:
                continue

def find_rightclick(pngname):
    # Ожидаем появление элемента
    sleep(0.5)
    r = None
    count = 200
    while r is None:
        count -= 1
        try:
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.8)
        except pyautogui.ImageNotFoundException:
            pass
        if count == 0:
            print(pngname + 'не найден')
            break
        else:
            if r != None:
                pyautogui.rightClick(r)
                return r
            else:
                continue

def find_tripleclick_rightclick(pngname):
    # Ожидаем появление элемента
    sleep(0.5)
    r = None
    count = 200
    while r is None:
        count -= 1
        try:
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.6)
        except pyautogui.ImageNotFoundException:
            pass
        if count == 0:
            print(pngname + 'не найден')
            break
        else:
            if r != None:
                pyautogui.tripleClick(r)
                pyautogui.rightClick(r)
                return r
            else:
                continue

def open_outlook():
    xoutlook, youtlook = pyautogui.locateCenterOnScreen(path + 'outlook.PNG', confidence=0.85)
    pyautogui.click(xoutlook, youtlook)
    youtlook -= 30
    pyautogui.click(xoutlook, youtlook)
    sleep(0.2)

def open_browser():
    x, y = pyautogui.locateCenterOnScreen(path + 'yndx.PNG', confidence=0.8)
    pyautogui.click(x, y)

def open_suid():
    x, y = pyautogui.locateCenterOnScreen(path + 'suid.PNG', confidence=0.8)
    pyautogui.click(x, y)

def open_sudir():
    try:
        x, y = pyautogui.locateCenterOnScreen(path + 'sudir.PNG', confidence=0.95)
        pyautogui.click(x, y)
    except pyautogui.ImageNotFoundException:
        pass


def mailmovetodone():
    pngname = 'moveto'
    find_click(pngname)
    pngname = 'donerequests'
    find_click(pngname)

def mailmovetoopublic():
    pngname = 'moveto'
    find_click(pngname)
    pngname = 'opublic'
    find_click(pngname)


def copyusername():
    xset, yset = pyautogui.locateCenterOnScreen(path + 'set.PNG', confidence=0.95)
    xset += 70
    pyautogui.moveTo(xset, yset)
    pyautogui.rightClick()
    n, m = pyautogui.locateCenterOnScreen(path + 'copymail.PNG', confidence=0.95)
    pyautogui.click(n, m)

def copyrolename():
    sleep(0.5)
    n, m = pyautogui.locateCenterOnScreen(path + 'rolename.PNG', confidence=0.7)
    m += 160
    n += 50
    pyautogui.tripleClick(n, m)
    pyautogui.rightClick(n, m)
    n, m = pyautogui.locateCenterOnScreen(path + 'copymail.PNG', confidence=0.95)
    pyautogui.click(n, m)

def request_suid():
    n, m = pyautogui.locateCenterOnScreen(path + 'requestsuid.PNG', confidence=0.95)
    pyautogui.click(n, m)

open_browser()
pngname = 'ksuit'
find_click(pngname)

pngname = 'more'
find_click(pngname)
pngname = 'masszakr'
find_click(pngname)
pngname = 'ksuitfield'
find_click(pngname)
pngname = 'zno'
find_click(pngname)
pngname = 'nextksuit'
find_click(pngname)
sleep(1)
pngname = 'nextksuit'
find_click(pngname)
