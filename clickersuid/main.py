from time import sleep

import pyautogui


# Указание количества запросов

requests = int(input("Количество запросов - "))
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True
# путь к файлам png

path = r'C:\Users\Andy\PycharmProjects\suidproject\png/'
# номер запроса
requestnum = 0

pngname = "Name"


def find_img(pngname):
    # Ожидаем появление элемента
    sleep(0.5)
    r = None
    count = 20
    while r is None:
        count -= 1
        r = pyautogui.locateCenterOnScreen(path + pngname, confidence=0.9)
        print(count)
        if count == 0:
            print(pngname + 'не найден')
            break
        else:
            if r != None:
                print(pngname + 'найден')
                return r
            else: continue

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


    # Если почта contractor доделать!!!
    try:
        n, m = pyautogui.locateCenterOnScreen(path + 'contractor1.PNG', confidence=0.95)
        pyautogui.click(n, m)
        sleep(0.5)
        n, m = pyautogui.locateCenterOnScreen(path + 'namescheck.PNG', confidence=0.95)
        pyautogui.click(n, m)
        n, m = pyautogui.locateCenterOnScreen(path + 'whom.PNG', confidence=0.95)
        n += 70
        pyautogui.doubleClick(n, m)
        n, m = pyautogui.locateCenterOnScreen(path + 'gpnmailblue.PNG', confidence=0.95)
        pyautogui.rightClick(n, m)
        n += 10
        m += 10
        pyautogui.click(n, m)
    except pyautogui.ImageNotFoundException:
            pass

    # open browser
    pyautogui.click(122, 1060)
    sleep(2)
    pyautogui.doubleClick(182, 308)
    # Ждем загр стр запроса
    sleep(6)
    pyautogui.click(207, 435)

    # paste
    sleep(1)
    pyautogui.rightClick(207, 435)
    pyautogui.click(239, 575)
    pyautogui.click(40, 462)
    # загрузка польз
    sleep(1)
    # название ресурса открыть Outlook
    open_outlook()
    copyrolename()

    # open browser
    pyautogui.click(122, 1060)

    # проверка жизни Выбор пользователей
    n, m = pyautogui.locateCenterOnScreen(path+'userchoose.PNG', confidence=0.95)

    # Выбрать пользователя и нажать далее
    sleep(4)

    # если пользователь не найден пробуем по УЗ
    try:
        n, m = pyautogui.locateCenterOnScreen(path+'nodata.PNG', confidence=0.95)
        # open outlook copy user account
        open_outlook()
        sleep(1)
        pyautogui.click(65, 1032)
        x, y = pyautogui.locateCenterOnScreen(path+'useraccount.PNG', confidence=0.95)
        y += 60
        pyautogui.tripleClick(x, y)
        pyautogui.rightClick()
        x, y = pyautogui.locateCenterOnScreen(path+'copymail.PNG', confidence=0.95)
        pyautogui.click(x, y)
        # open browser
        pyautogui.click(122, 1060)
        sleep(1)
        pyautogui.tripleClick(207, 435)

        pyautogui.rightClick(207, 435)
        x, y = pyautogui.locateCenterOnScreen(path+'paste.PNG', confidence=0.95)
        pyautogui.click(x, y)
        x, y = pyautogui.locateCenterOnScreen(path+'search.PNG', confidence=0.95)
        pyautogui.click(x, y)
        sleep(1)

        # open outlook
        open_outlook()
        # copy rolename
        n, m = pyautogui.locateCenterOnScreen(path + 'rolename.PNG', confidence=0.7)
        m += 160
        n += 50
        pyautogui.tripleClick(n, m)
        pyautogui.rightClick(n, m)

        # copy
        n, m = pyautogui.locateCenterOnScreen(path + 'copymail.PNG', confidence=0.95)
        pyautogui.click(n, m)

        # open browser
        pyautogui.click(122, 1060)
    except:
        pass

    # ищем куда нажать галочку
    try:
        ax, ay = pyautogui.locateCenterOnScreen(path+'checkbox1.PNG', grayscale=True, confidence=0.9)
        pyautogui.click(ax, ay)

        # next bx by
        bx, by = pyautogui.locateCenterOnScreen(path+'next.PNG', grayscale=True, confidence=0.9)
        pyautogui.click(bx, by)
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
    sleep(5)
    try:
        n, m = pyautogui.locateCenterOnScreen(path + 'circlecheckbox.PNG', confidence=0.95)
        pyautogui.click(n, m)
        bx, by = pyautogui.locateCenterOnScreen(path + 'next.PNG', grayscale=True, confidence=0.9)
        pyautogui.click(bx, by)
        sleep(2)
    except:
        pass

    # проверка жизни справки
    n, m = pyautogui.locateCenterOnScreen(path+'ident.PNG', grayscale=True, confidence=0.9)

    # Выбор инф ресурса
    pyautogui.doubleClick(279, 455)
    sleep(4)
    pyautogui.click(911, 502)
    pyautogui.moveTo(911, 520, 0.2)
    pyautogui.scroll(20000)
    pyautogui.click()
    pyautogui.click(151, 503)

    # проверка жизни Выбор доступа и перезапуск если не прогрузилась стр запроса
    n, m = pyautogui.locateCenterOnScreen(path+'dostupchoose.PNG', confidence=0.95)
    try:
        n, m = pyautogui.locateCenterOnScreen(path+'other.PNG', confidence=0.95)
    except pyautogui.ImageNotFoundException:
        continue

    # Вставить имя ресурса
    pyautogui.rightClick(154, 499)
    pyautogui.click(203, 647)
    sleep(1)

    # выбор типа ресурса
    pyautogui.click(1385, 506)
    pyautogui.moveTo(1351, 523, 0.1)
    pyautogui.scroll(2000)
    pyautogui.scroll(-300)
    sleep(1)
    n, m = pyautogui.locateCenterOnScreen(path+'infsys.PNG', grayscale=True, confidence=0.9)

    pyautogui.click(n, m)
    # search поиск нажать в инф ресурсах
    pyautogui.click(54, 563)
    # ждем загрузки ресурса
    sleep(5)

    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(path+'ident.PNG', grayscale=True, confidence=0.9)

    # choose выбираем
    sleep(2)
    pyautogui.click(87, 629)
    sleep(1)
    # next далее
    pyautogui.click(491, 368)
    sleep(1)

    # логика поиска крайнего срока
    pyautogui.click(259, 268)
    sleep(1)

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

    try:
        sleep(1)
        across, bcross = pyautogui.locateCenterOnScreen(path+'cross.PNG', confidence=0.7, region=(0, 0, 1269, 430))
        across += 50
        pyautogui.click(across, bcross)
    except pyautogui.ImageNotFoundException:
        pass

    # ищем 2023 2024 год
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

    try:
        sleep(1)
        pyautogui.locateCenterOnScreen(path+'inf.PNG', region=(250, 100, 1269, 350))
        pyautogui.click(1900, 368)
        pyautogui.click(491, 368)
    except pyautogui.ImageNotFoundException:
        pass

    # next
    pyautogui.doubleClick(491, 368)

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
    n, m = pyautogui.locateCenterOnScreen(path+'moveto.PNG', confidence=0.95)
    pyautogui.click(n, m)
    n, m = pyautogui.locateCenterOnScreen(path+'donerequests.PNG', confidence=0.95)
    pyautogui.click(n, m)

    print("Выполнен запрос № -", requestnum)