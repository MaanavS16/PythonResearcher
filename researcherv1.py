import keyboard
from pynput.mouse import Button, Controller as MouseController
from pynput.keyboard import Key, Controller as KeyboardController
import matplotlib.pyplot as plt
import time
import xlrd
import xlwt

workbook = xlrd.open_workbook('C:/Users/AnilSingh/dev/PCInput/SearchTerms.xlsx')
worksheet = workbook.sheet_by_name('Terms')
#print(worksheet.cell(0,0).value)
arrTerms = []
for i in range(worksheet.nrows):
    time.sleep(.05)
    if worksheet.cell(i, 0) == xlrd.empty_cell.value:
        #print("break")
        break
    else:
        arrTerms.append(worksheet.cell(i, 0).value)
        #print(worksheet.cell(i, 0).value)
#print(arrTerms)

def type(string):
    for char in string:
        keyboard.press(char)
        keyboard.release(char)
        #time.sleep(.12)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)


mouse = MouseController()
keyboard = KeyboardController()
arrPOI = [(710, 1052), (163, 55)]
num_terms = len(arrTerms)

mouse.position = (710, 1052)
mouse.click(Button.left, 1)
time.sleep(.2)

for i in range(num_terms):
    if i != 0:
        keyboard.press(Key.ctrl)
        keyboard.press("t")
        time.sleep(.1)
        keyboard.release("t")
        keyboard.release(Key.ctrl)
        time.sleep(.2)
    mouse.position = (163, 55)
    mouse.click(Button.left, 1)
    type(arrTerms[i])
    #print(i)

'''
arrSavedMouse = []
bln_run = True

while bln_run == True:
    if keyboard.is_pressed('q'):
        break
    print(mouse.position)
    arrSavedMouse.append(mouse.position)
'''
