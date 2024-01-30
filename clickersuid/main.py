from time import sleep

import pyautogui
import cv2


# pyautogui.position() alt+shift+e по запуску строки для определения пикселя
# Указание количества запросов

requests = int(input("Количество запросов - "))
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

requestnum = 0

# открыть СУДИР
try:
    # x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\sudir.PNG', confidence=0.95)
    pyautogui.click(72,1056)
except:
    pass



# Цикл

while requestnum < requests:
    # Обнуляем значения
    n, m, x, y, a, b, ax, ay, bx, by = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    requestnum += 1
    pyautogui.doubleClick(1917, 1055)

    # open browser
    pyautogui.click(122, 1060)
    # open suid
    pyautogui.click(44, 85)

    # ждем загрузки SUID и откр письмо

    # open outlook
    pyautogui.click(76, 1056)
    pyautogui.click(65, 1032)

    # copy username from mail
    pyautogui.click(439, 251)
    pyautogui.moveTo(519, 783, 0.5)
    pyautogui.rightClick()

    # copy
    pyautogui.click(530, 788)

    # Если почта contractor доделать
    try:
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\contractor1.PNG', confidence=0.95)
        pyautogui.click(n, m)
        sleep(0.5)
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\contractor2.PNG', confidence=0.95)
        pyautogui.doubleClick(n, m)
        pyautogui.rightClick()
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\clear.PNG', confidence=0.95)
        pyautogui.click(n, m)
        x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\dot.PNG', confidence=0.95)
        pyautogui.click(x, y)
        pyautogui.drag(-7, 0, 0.3, button='left')
        pyautogui.rightClick()
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\clear.PNG', confidence=0.95)
        pyautogui.click(n, m)
        pyautogui.tripleClick(x, y)
        pyautogui.rightClick()
        sleep(0.5)
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\copywin.PNG', confidence=0.95)
        pyautogui.click(n, m)
        pyautogui.position()
        pyautogui.click(1825, 650)
    except:
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
    # Название ресурса открыть Outlook
    pyautogui.click(76, 1056)
    pyautogui.click(65, 1032)

    # Открыть почту скопировать название ресурса
    sleep(1)
    pyautogui.tripleClick(1030, 787)
    pyautogui.rightClick()
    pyautogui.click(1090, 797)
    sleep(1)
    # open browser
    pyautogui.click(122, 1060)

    # проверка жизни Выбор пользователей
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\userchoose.PNG', confidence=0.95)

    # Выбрать пользователя и нажать далее
    sleep(3)



    # Если пользователь не найден пробуем по УЗ
    try:
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\nodata.PNG', confidence=0.95)
        # open outlook copy user account
        pyautogui.doubleClick(76, 1056)
        sleep(1)
        pyautogui.click(65, 1032)
        x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\useraccount.PNG', confidence=0.95)
        y += 60
        pyautogui.tripleClick(x, y)
        pyautogui.rightClick()
        x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\copymail.PNG', confidence=0.95)
        pyautogui.click(x, y)
        # open browser
        pyautogui.click(122, 1060)
        sleep(1)
        pyautogui.tripleClick(207, 435)

        pyautogui.rightClick(207, 435)
        x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\paste.PNG', confidence=0.95)
        pyautogui.click(x, y)
        x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\search.PNG', confidence=0.95)
        pyautogui.click(x, y)
        sleep(1)

        # open outlook
        pyautogui.click(76, 1056)
        pyautogui.click(65, 1032)

        pyautogui.tripleClick(1030, 787)
        pyautogui.rightClick()
        pyautogui.click(1090, 797)
        sleep(1)
        # open browser
        pyautogui.click(122, 1060)
    except:
        pass
    # ищем куда нажать галочку
    try:
        ax, ay = pyautogui.locateCenterOnScreen(r'C:\soft\png\checkbox1.PNG', grayscale=True, confidence=0.9)
        pyautogui.click(ax, ay)

        # next bx by
        bx, by = pyautogui.locateCenterOnScreen(r'C:\soft\png\next.PNG', grayscale=True, confidence=0.9)
        pyautogui.click(bx, by)
    except:
        pyautogui.locateCenterOnScreen(r'C:\soft\png\nodata.PNG', grayscale=True, confidence=0.9)
        print("Пользователь номер -", requestnum, "не найден");
        # open outlook
        pyautogui.click(76, 1056)
        pyautogui.click(65, 1032)
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\moveto.PNG', confidence=0.95)
        pyautogui.click(n, m)
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\manualrequest.PNG', confidence=0.95)
        pyautogui.click(n, m)

        continue



    # for multiple roles
    sleep(3)
    pyautogui.click(97, 454)
    pyautogui.click(498, 368)
    sleep(1)
    pyautogui.click(1003, 372)

    # проверка жизни справки
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\ident.PNG', grayscale=True, confidence=0.9)

    # Выбор инф ресурса
    pyautogui.doubleClick(279, 455)
    sleep(2)
    pyautogui.click(911, 502)
    pyautogui.moveTo(911, 520, 0.2)
    pyautogui.scroll(20000)
    pyautogui.click()
    pyautogui.click(151, 503)

    # проверка жизни Выбор доступа и перезапуск если не прогрузилась стр запроса
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\dostupchoose.PNG', confidence=0.95)
    try:
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\other.PNG', confidence=0.95)
    except pyautogui.ImageNotFoundException:
        continue

    # Вставить имя ресурса
    pyautogui.rightClick(154, 499)
    pyautogui.click(203, 647)
    sleep(1)

    # выбор типа ресурса
    pyautogui.click(1385, 506)
    pyautogui.moveTo(1351, 523, 0.2)
    pyautogui.scroll(2000)
    pyautogui.scroll(-300)
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\infsys.PNG', grayscale=True, confidence=0.9)

    pyautogui.click(n, m)
    # search поиск нажать в инф ресурсах
    pyautogui.click(54, 563)
    # ждем загрузки ресурса
    sleep(5)

    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\ident.PNG', grayscale=True, confidence=0.9)

    # choose выбираем
    pyautogui.click(87, 629)
    sleep(1)
    # next далее
    pyautogui.click(491, 368)
    sleep(1)

    # логика поиска крайнего срока
    pyautogui.click(259, 268)
    sleep(1)

    # проверка жизни Выбор срока действия
    try:
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\srokchoose.PNG', confidence=0.95)
    except pyautogui.ImageNotFoundException:
        print("Роль номер -",requestnum,"не найдена");
        # open outlook
        pyautogui.click(76, 1056)
        pyautogui.click(65, 1032)
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\moveto.PNG', confidence=0.95)
        pyautogui.click(n,m)
        n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\opublic.PNG', confidence=0.95)
        pyautogui.click(n, m)

        continue



    # Сдвинуть от крестика
    try:
        sleep(1)
        a, b = pyautogui.locateCenterOnScreen(r'C:\soft\png\cross.PNG', confidence=0.7, region=(0, 0, 1269, 430))
        a += 50
        pyautogui.click(a, b)
    except:
        pass


    # ищем 2023 2024 год или беск

    try:
        sleep(1)
        x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\date2023.PNG', grayscale=True, confidence=0.9,
                                              region=(250, 100, 1269, 350))
        pyautogui.doubleClick(x, y)
        sleep(1)
        c, d = pyautogui.locateCenterOnScreen(r'C:\soft\png\copy.PNG', confidence=0.7, region=(250, 100, 1269, 430))
        pyautogui.click(c, d)
        pyautogui.doubleClick(1637, 522)
        sleep(1)
        pyautogui.rightClick(1637, 522)
        pyautogui.click(1580, 660)
    except:
        pass

    try:
        x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\date3.PNG', grayscale=True, confidence=0.9,
                                              region=(250, 100, 1269, 350))
        sleep(1)
        pyautogui.doubleClick(x, y)
        sleep(2)
        c, d = pyautogui.locateCenterOnScreen(r'C:\soft\png\copy.PNG', confidence=0.7, region=(250, 100, 1269, 430))
        pyautogui.click(c, d)
        pyautogui.doubleClick(1637, 522)
        sleep(1)
        pyautogui.rightClick(1637, 522)
        pyautogui.click(1580, 660)
    except:
        pass

    try:
        x, y = pyautogui.locateCenterOnScreen(r'C:\soft\png\date4.PNG', grayscale=True, confidence=0.9,
                                              region=(250, 100, 1269, 350))
        sleep(1)
        pyautogui.doubleClick(x, y)
        sleep(2)
        c, d = pyautogui.locateCenterOnScreen(r'C:\soft\png\copy.PNG', confidence=0.7, region=(250, 100, 1269, 430))
        pyautogui.click(c, d)
        pyautogui.doubleClick(1637, 522)
        sleep(1)
        pyautogui.rightClick(1637, 522)
        pyautogui.click(1580, 660)
    except:
        pass

    try:
        sleep(1)
        pyautogui.locateCenterOnScreen(r'C:\soft\png\inf.PNG', region=(250, 100, 1269, 350))
        pyautogui.click(1900, 368)
        pyautogui.click(491, 368)
    except:
        pass

    # next
    pyautogui.doubleClick(491, 368)

    # open oneNote

    pyautogui.click(179, 1056)
    pyautogui.tripleClick(596, 264)
    pyautogui.rightClick()
    pyautogui.click(646, 302)

    # open browser

    pyautogui.click(122, 1060)
    # paste some words
    pyautogui.rightClick(334, 538)
    pyautogui.click(369, 682)

    # проверка жизни Обоснование
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\obosnovanie.PNG', confidence=0.95)

    # Завершить процесс для прода - нажать кнопку отправить запрос
    pyautogui.click(578, 371)
    # Переместить запрос в обработанные

    # open outlook
    pyautogui.click(76, 1056)
    pyautogui.click(65, 1032)
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\moveto.PNG', confidence=0.95)
    pyautogui.click(n, m)
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\donerequests.PNG', confidence=0.95)
    pyautogui.click(n, m)

    # Увеличить кол-во запросов для цикла

    print("Выполнен запрос № -", requestnum)