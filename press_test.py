import pyautogui
import time
import pandas as pd
import xlrd
import xlsxwriter
import numpy as np
import os


time.sleep(5)
pyautogui.hotkey('ctrl', 'o')
time.sleep(2)
pyautogui.click(pyautogui.locateCenterOnScreen('2_folder.png'))
pyautogui.click(pyautogui.locateCenterOnScreen('2_aaa.png'))
pyautogui.press(['down', 'down', 'down', 'down', 'down', 'down'])
pyautogui.press('enter')
time.sleep(2)
pyautogui.click(pyautogui.locateCenterOnScreen('3_view_edit.png'))
time.sleep(2)
pyautogui.click(pyautogui.locateCenterOnScreen('4_select.png'))
pyautogui.click()
distance = 2350
pyautogui.dragRel(distance, 0, duration=0.1)
pyautogui.click(pyautogui.locateCenterOnScreen('5_formant.png'))
pyautogui.press(['down', 'down', 'down', 'down', 'down',])
pyautogui.press('enter')
pyautogui.hotkey('ctrl', 's')
time.sleep(2)
pyautogui.click(pyautogui.locateCenterOnScreen('7_save.png'))
time.sleep(2)
pyautogui.click(pyautogui.locateCenterOnScreen('8_info.png'))
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('left')
pyautogui.press('enter')

i = 1
d = 1
path = 'C:/Users/Mateusz/Desktop/data/'
os.chdir(path)
Newfolder = 'Word' + str(i)
os.makedirs(Newfolder)
path2 = path + '\\' + Newfolder
os.chdir(path2)
df = pd.read_csv("C:/Users/Mateusz/Desktop/data/info.txt", delim_whitespace=True)
df.to_excel('C:/Users/Mateusz/Desktop/data/%s.xlsx' % i, 'Sheet1', index=False)
book = xlrd.open_workbook('C:/Users/Mateusz/Desktop/data/%s.xlsx' % i)
sheet = book.sheet_by_name('Sheet1')
txt_array = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
txt_array = np.array(txt_array)

max = len(txt_array)
while d < max:
    chunk = txt_array[d]
    chunk = np.delete(chunk, 0)
    chunk = chunk.reshape([4, 1])

    workbook = xlsxwriter.Workbook(path2 + '\\' + '%s.xlsx' % d)
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    for col, data in enumerate(chunk):
        worksheet.write_column(row, col, data)

    workbook.close()
    d += 1

