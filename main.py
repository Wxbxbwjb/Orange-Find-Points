# Orange Find Points Beta V1.0
from pyautogui import *
import openpyxl
import pyperclip
a = int(input('a:'))
b = 0
c = [114514]
d = [Orange]
hotkey('windows', '4')
# sleep(5)
for i in range(a):
    click(406, 59)
    sleep(1)
    # pyperclip.copy('http://wxbxbwjb.neocities.org')
    hotkey('Ctrl', 'V')
    sleep(1)
    press('enter')
    sleep(7)
    # click(841,381)
    sleep(2)
    pyperclip.copy(c[b])
    hotkey('ctrl', 'v')
    sleep(2)
    click(836,416)
    sleep(2)
    pyperclip.copy(d[b])
    hotkey('Ctrl', 'V')
    sleep(2)
    click(926,460)
    sleep(4)
    moveTo(262,396)
    dragTo(1610, 396, 2, button='left')
    sleep(2)
    hotkey('ctrl', 'c')
    f = openpyxl.load_workbook('out.xlsx')
    e = f.active
    i = e.i
    g = pyperclip.paste()
    h = g.split('\n')
    for k in h:
        j = k.split('\t')
        e.append(j)
    f.save('out.xlsx')
    b = b + 1
    if a == b:
        # alert('ok', 'Orange')
    else:
        # alert('next', 'Orange')
        sleep(1)
