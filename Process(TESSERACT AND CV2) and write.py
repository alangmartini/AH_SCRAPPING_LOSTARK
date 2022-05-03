# --------------------------------  TESSERACT e Insere no excel todas as informações

import pytesseract
import cv2
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



my_abas = ['gemchest', 'sailing', 'pets', 'mount', 
           'foraging', 'mining', 'hunting', 'fishing', 'excavating', 'traderother', 'logging']
my_sections =['name', 'avgprice', 'recentprices', 'lowestprice', 'cheapestrem']
my_sectionsletters = {'A', 'B', 'C', 'D', 'E'}

wb = load_workbook('auction.xlsx')


for aba in my_abas:
    excellinhafinal = 6
    ws = wb.active
    ws.title = f'{aba}'
    at = wb[f"{aba}"]
    if aba == 'pets':
        excellinhafinal = 8
    if aba == 'mount':
        excellinhafinal = 8
    if aba == 'gemchest':
        excellinhasfinal = 4
    if aba == 'foraging':
        excellinhafinal = 8
    if aba == 'hunting':
        excellinhafinal = 8
    if aba == 'fishing':
        excellinhafinal = 8
    if aba == 'mining':
        excellinhafinal = 5
    if aba == 'logging':
        excellinhafinal = 5
    if aba == 'excavating':
        excellinhafinal = 7
    if aba == 'traderother':
        excellinhafinal = 7

    for sections in my_sections:
        for linhas in range(2, excellinhafinal):
            #tesseract the images
            typing = (pytesseract.image_to_string(f'{aba}\worked{linhas}{sections}.bmp',  config= ' --psm 6 -c tessedit_char_whitelist={0123456789}'))
            if sections == 'name':
                typing = (pytesseract.image_to_string(f'{aba}\worked{linhas}{sections}.bmp',  config= ' --psm 6'))
            #make and retyping the tab
            if typing == '':
                typing = float(0)
            if sections == 'name':
                at[f'A{linhas}'] = typing
            if sections == 'avgprice':
                at[f'B{linhas}'] = float(typing)
            if sections == 'recentprices':
                at[f'C{linhas}'] =float(typing)
            if sections == 'lowestprice':
                at[f'D{linhas}'] = float(typing)
            if sections == 'cheapestrem':
                at[f'E{linhas}'] = float(typing)
            wb.save('auction.xlsx')
            
pyautogui.moveTo(0, 0)