import xml.etree.ElementTree as ET
from xml.dom import minidom

from os import listdir
import pyperclip
import pyautogui


def insert_item(item,x,y,duration):
    pyperclip.copy(item)
    pyautogui.doubleClick(x,y,duration=duration)
    pyautogui.hotkey('ctrl','v')

def open_exel():
    pyautogui.click(1919,1182,duration=0.5)
    pyautogui.doubleClick(270,39,duration=1)
    pyautogui.sleep(2)
    pyautogui.hotkey('ctrl','t')
    pyautogui.sleep(1)
    pyautogui.write('sheets.new',interval=0.1)
    pyautogui.press('enter')
    pyautogui.sleep(2)

def main():

    k = 0
    x = 0
    open_exel()
    pasta_notas = listdir("Notas")
    for nota in pasta_notas:
        x += 1
        k += 100
        y = 210

        xml = open("./Notas/"+nota)
        nfe = minidom.parse(xml)
        num_nfe = nfe.getElementsByTagName('nNF')

        insert_item(num_nfe[0].firstChild.data,k,y,duration= 2 if x == 1 else 0.2)
        y += 21

        itens = nfe.getElementsByTagName('det')
        for i in itens:
            num_item = i.attributes['nItem'].value
            insert_item(num_item,k,y,0.2)
            y += 21

            name_prod = i.getElementsByTagName('prod')[0].getElementsByTagName('xProd')[0].firstChild.data
            insert_item(name_prod,k,y,0.2)
            y += 21

            value_prod = i.getElementsByTagName('prod')[0].getElementsByTagName('vProd')[0].firstChild.data
            insert_item(value_prod,k,y,0.2)
            y += 21

        value_liq = nfe.getElementsByTagName('vNF')[0].firstChild.data
        insert_item(value_liq,k,y,0.2)
        y += 21
main()
