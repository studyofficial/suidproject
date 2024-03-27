from time import sleep

import pyautogui, sys

# Указание количества запросов
# requests = int(input("Количество запросов - "))
requests = 100
pyautogui.PAUSE = 0.5
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

open_sudir()




while requestnum < requests:

    # Обнуляем значения
    n, m, x, y, across, bcross, ax, ay, bx, by, xset, yset, xoutlook, youtlook = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    # Выводим время обработки запросов
    reqleft = requests - requestnum

    print("Выполняю запрос № -", requestnum)

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
        sleep(1)
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
        mailmovetoopublic()
        continue

    # multiple roles Если несколько ролей, выбрать одну (ВЫБОР ДОЛЖНОСТЕЙ)


    # проверка жизни справки
    pngname = 'ident'
    find_img(pngname)

    try:
        sleep(2.5)
        n, m = pyautogui.locateCenterOnScreen(path + 'circlecheckbox.PNG', confidence=0.95)
        pyautogui.click(n, m)
        n, m = pyautogui.locateCenterOnScreen(path + 'next.PNG', confidence=0.95)
        pyautogui.click(n, m)
    except:
        pass

    # Выбор инф ресурса (компания)
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
    find_click_exactly(pngname)
    try:
        n, m = pyautogui.locateCenterOnScreen(path+'scroll.PNG', confidence=0.95)
        pyautogui.click(n, m)
        pyautogui.scroll(2000)
        pyautogui.scroll(-300)
    except:
        pass

    pngname = 'infsys'
    find_click_exact(pngname)


    # search поиск нажать в инф ресурсах
    pngname = 'find'
    find_click(pngname)

    # choose выбираем по платформе PaaS

    try:
        pngname = 'dppaas'
        find_click(pngname)
        pngname = 'checkbox3'
        find_click(pngname)
    except:
        pyautogui.locateCenterOnScreen(path + 'nodata.PNG', grayscale=True, confidence=0.9)
        print("Пользователь номер -", requestnum, "не найден");
        # open outlook
        open_outlook()
        n, m = pyautogui.locateCenterOnScreen(path + 'moveto.PNG', confidence=0.95)
        pyautogui.click(n, m)
        n, m = pyautogui.locateCenterOnScreen(path + 'opublic.PNG', confidence=0.95)
        pyautogui.click(n, m)
        continue


    # next далее
    pngname = 'next'
    find_click_exact(pngname)

    # логика поиска крайнего срока
    pngname = 'chooseuser'
    find_click_kind(pngname)


    # проверка жизни выбор срока действия
    try:
        sleep(1)
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
    try:
        sleep(1)
        x, y = pyautogui.locateCenterOnScreen(path + 'cross.PNG', confidence=0.9)
        x += 50
        pyautogui.click(x, y)
    except pyautogui.ImageNotFoundException:
        pass
    # ищем 2023 2024 год

    try:
        sleep(1)
        pyautogui.locateCenterOnScreen(path+'inf.PNG', region=(250, 100, 1269, 350))
        pyautogui.click(1900, 368)
    except pyautogui.ImageNotFoundException:
        pass

    try:
        pyautogui.locateCenterOnScreen(path+'inf2.PNG', region=(250, 100, 1269, 350))
        pyautogui.click(1900, 368)
    except pyautogui.ImageNotFoundException:
        pass


    try:
        x, y = pyautogui.locateCenterOnScreen(path+'date4.PNG', grayscale=True, confidence=0.8,
                                              region=(250, 100, 1269, 350))
        sleep(1)
        pyautogui.doubleClick(x, y)
        sleep(2)
        c, d = pyautogui.locateCenterOnScreen(path+'copy.PNG', confidence=0.7, region=(250, 100, 1269, 430))
        pyautogui.click(c, d)
        pngname = 'lastday'
        find_tripleclick_rightclick(pngname)
        pngname = 'paste'
        find_click(pngname)
    except pyautogui.ImageNotFoundException:
        pass



    pngname = 'next'
    find_click_exact(pngname)

    # open oneNote
    pngname = 'onenote'
    find_click(pngname)
    pngname = 'obosnov'
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


    # нажать кнопку отправить запрос
    pngname = 'end'
    find_click_exactly(pngname)

    pngname = 'end2'
    find_click_exactly(pngname)
    # Переместить запрос в обработанные

    # open outlook
    open_outlook()
    mailmovetodone()
