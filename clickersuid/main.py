from time import sleep

import pyautogui

# Указание количества запросов
requests = int(input("Количество запросов - "))
pyautogui.PAUSE = 0.4
pyautogui.FAILSAFE = True

# путь к файлам png
path = r'C:\Users\Andy\PycharmProjects\suidproject\png/'
# номер запроса
requestnum = 0

pngname = "Name"


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
            r = pyautogui.locateCenterOnScreen(path + pngname + '.PNG', confidence=0.8)
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


def mailmoveto():
    n, m = pyautogui.locateCenterOnScreen(path+'moveto.PNG', confidence=0.95)
    pyautogui.click(n, m)
    n, m = pyautogui.locateCenterOnScreen(path+'donerequests.PNG', confidence=0.95)
    pyautogui.click(n, m)


def copyusername():
    xset, yset = pyautogui.locateCenterOnScreen(path + 'set.PNG', confidence=0.95)
    xset += 70
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
    n, m = pyautogui.locateCenterOnScreen(path + 'requestsuid.PNG', confidence=0.95)
    pyautogui.click(n, m)

open_sudir()




while requestnum < requests:

    # Обнуляем значения
    n, m, x, y, across, bcross, ax, ay, bx, by, xset, yset, xoutlook, youtlook = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    # Выводим время обработки запросов
    reqleft = requests - requestnum

    print("Осталось запросов -", reqleft, ". Это займет примерно ", reqleft, "мин.")

    requestnum += 1

    open_browser()
    open_suid()
    open_outlook()

    # copy username from mail
    copyusername()

    # open browser
    open_browser()
    pngname = 'requestsuid'
    find_click(pngname)
    find_click(pngname)
    # Ждем загр стр запроса
    pngname = 'emptyslot'
    find_rightclick(pngname)
    pngname = 'paste'
    find_click(pngname)
    pngname = 'search'
    find_click(pngname)
    # название ресурса открыть Outlook
    open_outlook()
    copyrolename()

    open_browser()

    # проверка жизни Выбор пользователей
    pngname = 'userchoose'
    find_img(pngname)

    # Выбрать пользователя и нажать далее


    # если пользователь не найден пробуем по УЗ
    try:
        sleep(2)
        x, y = pyautogui.locateCenterOnScreen(path + 'nodata.PNG', confidence=0.95)

        open_outlook()

        x, y = pyautogui.locateCenterOnScreen(path+'useraccount.PNG', confidence=0.95)
        y += 60
        pyautogui.tripleClick(x, y)
        pyautogui.rightClick()
        pngname = 'copymail'
        find_click(pngname)
        # open browser
        open_browser()
        pngname = 'emptyslot'
        find_tripleclick_rightclick(pngname)

        pngname = 'paste'
        find_click(pngname)

        pngname = 'search'
        find_click(pngname)

        open_outlook()
        # copy rolename
        n, m = pyautogui.locateCenterOnScreen(path + 'rolename.PNG', confidence=0.7)
        m += 160
        n += 50
        pyautogui.tripleClick(n, m)
        pyautogui.rightClick(n, m)

        # copy
        pngname = 'copymail'
        find_click(pngname)

        open_browser()
    except:
        pass

    # ищем куда нажать галочку
    try:
        pngname = 'checkbox1'
        find_click(pngname)

        pngname = 'next'
        find_click_exact(pngname)
    except:
        pyautogui.locateCenterOnScreen(path+'nodata.PNG', grayscale=True, confidence=0.9)
        print("Пользователь номер -", requestnum, "не найден");
        # open outlook
        open_outlook()
        n, m = pyautogui.locateCenterOnScreen(path+'moveto.PNG', confidence=0.95)
        pyautogui.click(n, m)
        n, m = pyautogui.locateCenterOnScreen(path+'manualrequest.PNG', confidence=0.95)
        pyautogui.click(n, m)

        continue

    # multiple roles Если несколько ролей, выбрать одну (ВЫБОР ДОЛЖНОСТЕЙ)
    try:
        pngname = 'circlecheckbox'
        find_click(pngname)

        pngname = 'next'
        find_click_exact(pngname)
    except:
        pass

    # проверка жизни справки
    pngname = 'ident'
    find_img(pngname)

    # Выбор инф ресурса
    pngname = 'another'
    find_click(pngname)

    pngname = 'company'
    find_click_kind(pngname)
    pngname = 'scroll'
    find_click(pngname)

    pyautogui.scroll(20000)
    pngname = 'emptyscroll'
    find_click_kind(pngname)

    pngname = 'dostupchoose'
    find_img(pngname)

    # Вставить имя ресурса
    pngname = 'nameres'
    find_rightclick(pngname)
    pngname = 'paste'
    find_click(pngname)

    # выбор типа ресурса
    pngname = 'typeres'
    find_click_exact(pngname)
    pngname = 'scroll'
    find_click(pngname)
    pyautogui.scroll(2000)
    pyautogui.scroll(-300)
    pngname = 'infsys'
    find_click_exact(pngname)

    # search поиск нажать в инф ресурсах
    pngname = 'find'
    find_click(pngname)

    # choose выбираем
    pngname = 'checkbox3'
    find_click(pngname)

    # next далее
    pngname = 'next'
    find_click_exact(pngname)

    # логика поиска крайнего срока
    pngname = 'chooseuser'
    find_click_kind(pngname)


    # проверка жизни выбор срока действия
    try:
        n, m = pyautogui.locateCenterOnScreen(path+'srokchoose.PNG', confidence=0.95)
    except pyautogui.ImageNotFoundException:
        print("Роль номер -",requestnum,"не найдена")
        # open outlook
        open_outlook()
        n, m = pyautogui.locateCenterOnScreen(path+'moveto.PNG', confidence=0.95)
        pyautogui.click(n,m)
        n, m = pyautogui.locateCenterOnScreen(path+'opublic.PNG', confidence=0.95)
        pyautogui.click(n, m)

        continue

    # поиск крестика
    pngname = 'cross'
    find_click_kind(pngname)

    # ищем 2023 2024 год

    try:
        sleep(1)
        pyautogui.locateCenterOnScreen(path+'inf.PNG', region=(250, 100, 1269, 350))
        pyautogui.click(1900, 368)
    except pyautogui.ImageNotFoundException:
        pass

    try:
        sleep(1)
        x, y = pyautogui.locateCenterOnScreen(path+'date2023.PNG', grayscale=True, confidence=0.9,
                                              region=(250, 100, 1269, 350))
        pyautogui.doubleClick(x, y)
        sleep(1)
        c, d = pyautogui.locateCenterOnScreen(path+'copy.PNG', confidence=0.7, region=(250, 100, 1269, 430))
        pyautogui.click(c, d)
        pyautogui.doubleClick(1637, 522)
        sleep(1)
        pyautogui.rightClick(1637, 522)
        pyautogui.click(1580, 660)
    except pyautogui.ImageNotFoundException:
        pass

    try:
        x, y = pyautogui.locateCenterOnScreen(path+'date3.PNG', grayscale=True, confidence=0.9,
                                              region=(250, 100, 1269, 350))
        sleep(1)
        pyautogui.doubleClick(x, y)
        sleep(2)
        c, d = pyautogui.locateCenterOnScreen(path+'copy.PNG', confidence=0.7, region=(250, 100, 1269, 430))
        pyautogui.click(c, d)
        pyautogui.doubleClick(1637, 522)
        sleep(1)
        pyautogui.rightClick(1637, 522)
        pyautogui.click(1580, 660)
    except pyautogui.ImageNotFoundException:
        pass

    try:
        x, y = pyautogui.locateCenterOnScreen(path+'date4.PNG', grayscale=True, confidence=0.9,
                                              region=(250, 100, 1269, 350))
        sleep(1)
        pyautogui.doubleClick(x, y)
        sleep(2)
        c, d = pyautogui.locateCenterOnScreen(path+'copy.PNG', confidence=0.7, region=(250, 100, 1269, 430))
        pyautogui.click(c, d)
        pyautogui.doubleClick(1637, 522)
        sleep(1)
        pyautogui.rightClick(1637, 522)
        pyautogui.click(1580, 660)
    except pyautogui.ImageNotFoundException:
        pass



    # next
    pngname = 'next'
    find_click_exact(pngname)

    # open oneNote

    pyautogui.click(179, 1056)
    pyautogui.tripleClick(596, 264)
    pyautogui.rightClick()
    pyautogui.click(646, 302)

    # open browser
    open_browser()
    # paste some words
    pyautogui.rightClick(334, 538)
    pyautogui.click(369, 682)

    # проверка жизни Обоснование
    n, m = pyautogui.locateCenterOnScreen(path+'obosnovanie.PNG', confidence=0.95)

    # нажать кнопку отправить запрос
    pyautogui.click(578, 371)
    # Переместить запрос в обработанные

    # open outlook
    open_outlook()
    mailmoveto()

    print("Выполнен запрос № -", requestnum)