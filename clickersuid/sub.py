from time import sleep

import pyautogui
import cv2


# pyautogui.position() alt+shift+e по запуску строки для определения пикселя
# Указание количества запросов

requests = int(input("Количество запросов - "))
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True

requestnum = 0

# Цикл

while requestnum < requests:
    # Обнуляем значения
    n, m, x, y, a, b, ax, ay, bx, by, cx, cy = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0



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
    sleep(1)
    # copy
    pyautogui.click(530, 788)
    # open browser
    pyautogui.click(122, 1060)
    sleep(8)
    pyautogui.click(192, 189)
    sleep(1)
    pyautogui.click(192, 209)
    sleep(2)
    pyautogui.click(554, 311)
    sleep(3)
    # Ждем загр стр запроса

    pyautogui.click(207, 455)
    # paste
    sleep(1)
    pyautogui.rightClick(207, 455)
    pyautogui.click(239, 605)
    pyautogui.click(40, 482)
    # загрузка польз
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

    # choose ax ay
    ax, ay = pyautogui.locateCenterOnScreen(r'C:\soft\png\checkbox1.PNG', grayscale=True, confidence=0.9)
    pyautogui.click(ax, ay)
    # next bx by
    bx, by = pyautogui.locateCenterOnScreen(r'C:\soft\png\next.PNG', grayscale=True, confidence=0.9)
    pyautogui.click(bx, by)


    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\ident.PNG', grayscale=True, confidence=0.9)
    # конец проверки

    # Выбор инф ресурса
    sleep(2)
    # Вставить имя ресурса
    pyautogui.rightClick(339, 456)
    pyautogui.click(371, 605)
    sleep(1)


    # search поиск нажать в инф ресурсах
    pyautogui.click(41, 505)
    sleep(4)

    pyautogui.position()
    cx, cy = pyautogui.locateCenterOnScreen(r'C:\soft\png\checkbox2.PNG', confidence=0.99)
    pyautogui.click(cx, cy)
    # ждем загрузки ресурса
    sleep(3)

    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\ident.PNG', grayscale=True, confidence=0.9)
    # конец проверки

    # next далее
    pyautogui.click(bx, by)

    # несколько доступов
    try:
        sleep(1)
        ax, ay = pyautogui.locateCenterOnScreen(r'C:\soft\png\checkbox2.PNG', grayscale=True, confidence=0.9)
        pyautogui.click(ax, ay)
    except:
        pass

    # проверка жизни
    n, m = pyautogui.locateCenterOnScreen(r'C:\soft\png\ident.PNG', grayscale=True, confidence=0.9)
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
    pyautogui.click(76, 1056)
    pyautogui.click(65, 1032)

    # drag mail
    pyautogui.moveTo(404, 253)
    pyautogui.dragTo(250, 249, 0.5, button='left')

    # Увеличить кол-во запросов для цикла
    requestnum += 1
