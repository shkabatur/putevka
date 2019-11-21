# -*- coding: utf-8 -*-

from openpyxl import Workbook, load_workbook
from pprint import pprint
from fpdf import FPDF
import datetime
import json
import re

with open("position.json") as json_file:
    position = json.load(json_file)


pdf = FPDF(orientation="L", unit="mm", format="A4")
pdf.add_font("KEK", '', 'times-new-roman.ttf', uni=True)

DATE = "D"
ADDRESS = "K"
NAME = "C"
RODITEL = "M"
MINISTERSTVO = "Q"

SMENA = "13"

s_god = "9"
po_god = "9"

s_den = "11"
po_den = "23"

s_mesyac = "июнь"
po_mesyac = "июль"

workbook = load_workbook(filename="Astrakhan.xlsx")
sheet = workbook.active
kinds= []

def getAges(a,b):
    b = datetime.datetime.now()
    return str(int((b - a).days / 365))    


def print_xy(x,y,text):
    pdf.set_xy(x,y)
    pdf.cell(0,0,text)

def printInCells(s, x, y):
    INTERVAL = 6.2
    for c in s :
        print_xy(x,y,c)
        x += INTERVAL


def printKind(kind):
    pdf.add_page()
    pdf.set_font("KEK", size=14)
    #==========================================================================
    #========Первая страничка==================================================
    #nomer smeny
    print_xy(position["1s"]["nomer_smeny"]["x"], position["1s"]["nomer_smeny"]["y"], SMENA)

    #S DATE
    print_xy(position["1s"]["s_den"]["x"], position["1s"]["s_den"]["y"], s_den)
    print_xy(position["1s"]["s_mesyac"]["x"], position["1s"]["s_mesyac"]["y"], s_mesyac)
    print_xy(position["1s"]["s_god"]["x"], position["1s"]["s_god"]["y"], s_god)

    #PO DATE
    print_xy(position["1s"]["po_den"]["x"], position["1s"]["po_den"]["y"], po_den)
    print_xy(position["1s"]["po_mesyac"]["x"], position["1s"]["po_mesyac"]["y"], po_mesyac)
    print_xy(position["1s"]["po_god"]["x"], position["1s"]["po_god"]["y"], po_god)

    #VOZRAST
    print_xy(position["1s"]["vozrast"]["x"], position["1s"]["vozrast"]["y"], kind["ages"])

    #FAMILIYA
    printInCells(kind["first_name"], position["1s"]["Familiya"]["x"], position["1s"]["Familiya"]["y"])
    #IMYA
    printInCells(kind["last_name"], position["1s"]["Imya"]["x"], position["1s"]["Imya"]["y"])
    #OTCHESTVO
    printInCells(kind["patronymic"], position["1s"]["Otchestvo"]["x"], position["1s"]["Otchestvo"]["y"])

    #FIO_RODITELYA
    print_xy(position["1s"]["FIO_roditelya"]["x"],position["1s"]["FIO_roditelya"]["y"], kind["parent"][0])
    if len(kind["parent"]) > 1:
        print_xy(position["1s"]["FIO_roditelya"]["x"],position["1s"]["FIO_roditelya"]["y"] + 7, kind["parent"][1])


    #Adres_roditelya
    print_xy(position["1s"]["Adres_roditelya"]["x"],position["1s"]["Adres_roditelya"]["y"], kind["address"][0])
    if len(kind["address"]) > 1:
        print_xy(position["1s"]["Adres_roditelya"]["x"]-35,position["1s"]["Adres_roditelya"]["y"]+7, kind["address"][1])

    #Ministerstvo
    print_xy(position["1s"]["ministerstvo"]["x"],position["1s"]["ministerstvo"]["y"], kind["ministerstvo"])
    #LINE
    #x = position["1s"]["line"]["x"]
    #y = position["1s"]["line"]["y"]
    #pdf.set_line_width(2)
    #pdf.line(x,y,x+35,y)
    #===================Конец первой странички==================================
    #===========================================================================


i = 10
while sheet[NAME + str(i)].value:
    kind = {}
    kind["first_name"], kind["last_name"], kind["patronymic"] = sheet[NAME + str(i)].value.split()
    kind["date"] = sheet[DATE + str(i)].value
    kind["parent"] = re.split("\n|,|  ", sheet[RODITEL + str(i)].value)
    kind["ministerstvo"] = sheet[MINISTERSTVO + str(i)].value
    address = sheet[ADDRESS + str(i)].value
    if len(address) > 40:
        sp = address.split(',')
        kind["address"] = [",".join(sp[:(len(sp)//2)]),",".join(sp[(len(sp)//2):])]
    else:
        kind["address"] = [address]
    kind["ages"] = kind["date"].strftime("%d.%m.%Y") + " (" + getAges(kind["date"],datetime.datetime.now()) + " лет)"
    kinds.append(kind)
    i += 1

def main():
    for kind in kinds:
        printKind(kind)
    pdf.output("kek.pdf")


if __name__ == "__main__":
    main()
