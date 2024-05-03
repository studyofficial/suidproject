from time import sleep

import pyautogui, sys

# Указание количества запросов
requests = 100
pyautogui.PAUSE = 1.5
pyautogui.FAILSAFE = True
path = r'C:\Users\Andy\PycharmProjects\suidproject\png/'

requestnum = 0


def find_img(pngname):
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
    count = 200
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
                pyautogui.click(r)
                return r
            else:
                continue

def find_click_fast(pngname):
    # Ожидаем появление элемента
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
    pngname = 'donewithdraw'
    find_click(pngname)

def mailmovetomanual():
    pngname = 'moveto'
    find_click(pngname)
    pngname = 'manualwithdraw'
    find_click(pngname)


def copyusername():
    xset, yset = pyautogui.locateCenterOnScreen(path + 'delete.PNG', confidence=0.95)
    xset += 90
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
    sleep(2)
    pngname = 'zayavki'
    find_click(pngname)
    pngname = 'plus'
    find_click(pngname)



open_sudir()

pyautogui.click(1917,1055)

# Цикл

while requestnum < requests:
    # Обнуляем значения
    n, m, x, y, across, bcross, ax, ay, bx, by, xset, yset, xoutlook, youtlook = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    print('Выполняю запрос №',requestnum)
    open_browser()
    open_suid()
    open_outlook()

    copyusername()

    # open browser
    open_browser()
    request_suid()
    pngname = 'withdraw'
    find_click(pngname)
    sleep(2)
    pngname = 'emptyslot'
    find_rightclick(pngname)


    pngname = 'paste'
    find_click(pngname)

    pngname = 'find'
    find_click(pngname)
    # загрузка польз

    open_outlook()



    # Открыть почту скопировать название ресурса
    copyrolename()
    sleep(1)
    # open browser
    open_browser()

    try:
        pngname = 'checkbox1'
        find_click(pngname)

        pngname = 'next'
        find_click_exact(pngname)
    except:
        pyautogui.locateCenterOnScreen(path + 'nodata.PNG', grayscale=True, confidence=0.9)
        print("Пользователь номер -", requestnum, "не найден");
        # open outlook
        open_outlook()
        mailmovetomanual()
        requestnum += 1
        continue


    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(path+'ident.PNG', grayscale=True, confidence=0.9)
    # конец проверки

    # Выбор инф ресурса
    sleep(3)
    # Вставить имя ресурса
    pngname = 'dostupname'
    find_rightclick(pngname)
    pngname = 'paste'
    find_click(pngname)
    sleep(1)


    # search поиск нажать в инф ресурсах
    pngname = 'find'
    find_click(pngname)
    sleep(2)

    pngname = 'checkbox2'
    find_click(pngname)

    try:
        pyautogui.locateCenterOnScreen(path + 'nodata2.PNG', grayscale=True, confidence=0.9)
        print("Пользователь номер -", requestnum, "не найден");
        # open outlook
        open_outlook()
        mailmovetomanual()
        requestnum += 1
    except:
        pass


    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(path+'ident.PNG', grayscale=True, confidence=0.9)
    # конец проверки

    # next далее
    pngname = 'next'
    find_click(pngname)

    try:
        pngname = 'checkbox2'
        find_click_fast(pngname)
    except:
            pass

    pngname = 'next'
    find_click(pngname)


    # open oneNote
    pngname = 'onenote'
    find_click(pngname)
    pngname = 'obosnovwithdraw'
    find_tripleclick_rightclick(pngname)
    pngname = 'copyonenote'
    find_click(pngname)
    # open browser

    open_browser()
    # paste some words
    pngname = 'obosnovfield'
    find_rightclick(pngname)
    pngname = 'paste'
    find_click(pngname)

    # Завершить процесс для прода - нажать кнопку отправить запрос
    pngname = 'end'
    find_click_exact(pngname)

    # Переместить запрос в обработанные

    # open outlook
    open_outlook()
    # drag mail
    mailmovetodone()

    # Увеличить кол-во запросов для цикла
    requestnum += 1
