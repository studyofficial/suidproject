from time import sleep

import pyautogui

# Указание количества запросов
requests = int(input("Количество запросов - "))
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True
path = r'C:\Users\Andy\PycharmProjects\suidproject\png/'

requestnum = 0

def open_outlook():
    xoutlook, youtlook = pyautogui.locateCenterOnScreen(path + 'outlook.PNG', confidence=0.85)
    pyautogui.click(xoutlook, youtlook)
    youtlook -= 30
    pyautogui.click(xoutlook, youtlook)
    sleep(0.2)
    print('Outlook открыт')

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

def copymail():
    n, m = pyautogui.locateCenterOnScreen(path + 'copymail.PNG', confidence=0.95)
    pyautogui.click(n, m)

def copyusername():
    xset, yset = pyautogui.locateCenterOnScreen(path + 'delete.PNG', confidence=0.95)
    xset += 90
    pyautogui.moveTo(xset, yset)
    pyautogui.rightClick()
    n, m = pyautogui.locateCenterOnScreen(path + 'copymail.PNG', confidence=0.95)
    pyautogui.click(n, m)

def copyrolename():
    n, m = pyautogui.locateCenterOnScreen(path + 'rolename.PNG', confidence=0.7)
    m += 160
    n += 50
    pyautogui.tripleClick(n, m)
    pyautogui.rightClick(n, m)
    n, m = pyautogui.locateCenterOnScreen(path + 'copymail.PNG', confidence=0.95)
    pyautogui.click(n, m)

def request_suid():
    n, m = pyautogui.locateCenterOnScreen(path + 'zayavki.PNG', confidence=0.95)
    pyautogui.click(n, m)
    m += 25
    pyautogui.click(n, m)


open_sudir()


# Цикл

while requestnum < requests:
    # Обнуляем значения
    n, m, x, y, across, bcross, ax, ay, bx, by, xset, yset, xoutlook, youtlook = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    open_browser()
    open_suid()
    open_outlook()

    copyusername()

    # open browser
    open_browser()
    sleep(5)
    request_suid()
    sleep(3)
    pyautogui.click(554, 311)
    sleep(3)


    pyautogui.click(207, 455)
    # paste
    sleep(2)
    pyautogui.rightClick(207, 455)
    pyautogui.click(239, 605)
    pyautogui.click(40, 482)
    # загрузка польз

    open_outlook()



    # Открыть почту скопировать название ресурса
    sleep(1)
    pyautogui.tripleClick(1030, 787)
    pyautogui.rightClick()
    pyautogui.click(1090, 797)
    sleep(1)
    # open browser
    pyautogui.click(122, 1060)

    # choose ax ay
    ax, ay = pyautogui.locateCenterOnScreen(path+'checkbox1.PNG', grayscale=True, confidence=0.9)
    pyautogui.click(ax, ay)
    # next bx by
    bx, by = pyautogui.locateCenterOnScreen(path+'next.PNG', grayscale=True, confidence=0.9)
    pyautogui.click(bx, by)


    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(path+'ident.PNG', grayscale=True, confidence=0.9)
    # конец проверки

    # Выбор инф ресурса
    sleep(3)
    # Вставить имя ресурса
    pyautogui.rightClick(339, 456)
    pyautogui.click(371, 605)
    sleep(1)


    # search поиск нажать в инф ресурсах
    pyautogui.click(41, 505)
    sleep(4)

    pyautogui.position()
    cx, cy = pyautogui.locateCenterOnScreen(path+'checkbox2.PNG', confidence=0.99)
    pyautogui.click(cx, cy)
    # ждем загрузки ресурса
    sleep(3)

    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(path+'ident.PNG', grayscale=True, confidence=0.9)
    # конец проверки

    # next далее
    pyautogui.click(bx, by)

    # несколько доступов
    try:
        sleep(1)
        ax, ay = pyautogui.locateCenterOnScreen(path+'checkbox2.PNG', grayscale=True, confidence=0.9)
        pyautogui.click(ax, ay)
    except:
        pass

    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(path+'ident.PNG', grayscale=True, confidence=0.9)
    # конец проверки

    # next далее
    pyautogui.click(bx, by)

    # open oneNote

    pyautogui.click(179, 1056)
    pyautogui.position()
    pyautogui.tripleClick(645, 686)
    pyautogui.rightClick()
    pyautogui.click(691, 725)

    # open browser

    pyautogui.click(122, 1060)
    # paste some words
    pyautogui.rightClick(334, 538)
    pyautogui.click(369, 682)

    # Завершить процесс для прода - нажать кнопку отправить запрос
    pyautogui.click(578, 371)
    # Переместить запрос в обработанные

    # open outlook
    xoutlook, youtlook = pyautogui.locateCenterOnScreen(path + 'outlook.PNG', confidence=0.85)
    pyautogui.click(xoutlook, youtlook)
    youtlook -= 30
    pyautogui.click(xoutlook, youtlook)
    sleep(0.5)

    # drag mail
    n, m = pyautogui.locateCenterOnScreen(path + 'moveto.PNG', confidence=0.95)
    pyautogui.click(n, m)
    n, m = pyautogui.locateCenterOnScreen(path + 'donewithdraw.PNG', confidence=0.95)
    pyautogui.click(n, m)

    # Увеличить кол-во запросов для цикла
    requestnum += 1
