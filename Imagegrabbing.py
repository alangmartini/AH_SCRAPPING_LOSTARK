'''
Thats a program that alt tabs to market in game LOST ARK, click on the respective
tabs, and, for every line of product in the AH (auction house), it takes five
screenshots, one for each row: name, avg.price, recent prices, lowest prices, and cheapest.rem.

After that, the screenshots goes through a CV2 processing, and then is read with TESSERACT,
for translating these images in integers. They are then inserted in a excel file (auction.xlsx)
and reorganized in a single sheet in a second excel file (auctionresults.xlsx)


How to use:
Set your game to 1080x1024 with forced 21:9. And in the upper left side of your screen.
Put your windows bar to the left as well. Honestly, this should've been made with Client coordinates, and not screen.

Take a screenshot of every tab you want to search, and cut it with only the name
eg: Gemchest, Pets, Mounts.

This currently does not support multiple tabs. Since i  made for my own use, and i
only cared about about gathering and ship parts prices.
'''





from PIL import Image
import pandas as pd
from openpyxl import Workbook, load_workbook
import pyscreenshot as ImageGrab
import pyautogui
import time
import pygetwindow
import numpy as np

pyautogui.FAILSAFE= False
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

#---------------------------------------------read all market #print
#screen iten names
pyautogui.hotkey('alt', 'tab') # alt tab to change to the game. Modify it later to winactivate
time.sleep(1)
#print everything
my_list = ['other', 'gemchest', 'mount', 'pets', 'sailing'] #tabs  that you want to screen
my_numbertags =[ 'avgprice', 'recentprices', 'lowestprice']
my_bigtags = ['name', 'cheapestrem']

for aba in my_list:
    # variables that need to be defined for every 'abas' run
    excelline1 = 2
    excelline2 = 8
    coordx1 = 691
    coordx2 = 900
    yn1 = 336
    yn2 = 377
    #click on the tab
    click = pyautogui.locateOnScreen(f'{aba}.bmp')
    time.sleep(0.1)
    pyautogui.moveTo(click)
    time.sleep(0.2)
    pyautogui.click()
    time.sleep(2)
    #click on the "all" section
    clickall = pyautogui.locateOnScreen('all.bmp')
    time.sleep(0.1)
    pyautogui.moveTo(clickall)
    time.sleep(0.2)
    pyautogui.click()
    time.sleep(2)
    # define different amount of lines for a few cases. 
    # Since the time difference in processing 3 lines of market a
    # and 10 is very big, this helps saving some time
    if aba == 'mount':
        excelline2 = 8
    if aba == 'pets':
        excelline2 = 8 
    if aba == 'other':
        excelline2 = 3
    if aba == 'gemchest':
         excelline2 = 4
    if aba == 'sailing':
         excelline2 = 6
    #imagegrabbing name related tabs (which have a bigger size)
    for sections in my_bigtags:     
        for unidades in range(excelline1, excelline2): #x for the name rank. Yn is common
            im = ImageGrab.grab(bbox=(coordx1, yn1, coordx2, yn2))
            im.save(f'{aba}\ {unidades}{sections}.bmp')
            yn1 = yn1 + 56 # the beginning of the second iten is +56y from the first
            yn2 = yn2 + 56 # as do the end       
        coordx1 = coordx1 + 868
        coordx2 = coordx2 + 766
        yn1 = 336
        yn2 = 377
    #defining variables and imagegrabbing moneyrelated tabs
    #which are much smaller.
    moneyyn1 = 344
    moneyyn2 = 368
    moneyxn1 = 1019 
    moneyxn2 = 1105
    for sections in my_numbertags:     
        for unidades in range(excelline1, excelline2):
                im = ImageGrab.grab(bbox=(moneyxn1, moneyyn1, moneyxn2, moneyyn2))
                im.save(f'{aba}\ {unidades}{sections}.bmp')
                moneyyn1 = moneyyn1 + 56 # the beginning of the second iten is +53y from the first
                moneyyn2 = moneyyn2 + 56 # as do the end     
        moneyxn1 = moneyxn1 + 160
        moneyxn2 = moneyxn2 + 154
        moneyyn1 = 344
        moneyyn2 = 368
    pyautogui.moveTo(click)
    time.sleep(1)
    pyautogui.click()
    time.sleep(2)

    
    