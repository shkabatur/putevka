# -*- coding: utf-8 -*-

from openpyxl import Workbook, load_workbook
from pprint import pprint
from fpdf import FPDF
import datetime

time_new_roman = 'times-new-roman.ttf'

DATE = "D"
ADDRESS = "K"
NAME = "C"
MAMA = "M"

def getAges(a):
    b = datetime.datetime.now()
    return int((b - a).days / 365)    

workbook = load_workbook(filename="Astrakhan.xlsx")
sheet = workbook.active
kinds= []

i = 10
while sheet[NAME + str(i)].value:
    kind = {}
    kind["name"] = sheet[NAME + str(i)].value
    kind["date"] = sheet[DATE + str(i)].value
    kind["mama"] = sheet[MAMA + str(i)].value
    kind["address"] = sheet[ADDRESS + str(i)].value
    kind["ages"] = getAges(kind["date"])
    kinds.append(kind)
    i += 1

#pprint(kinds)
#-----------------------------------------------
#-------GENERATING PDF FILES
#-----------------------------------------------

pdf = FPDF(orientation="L", unit="mm", format="A4")
pdf.add_font("KEK", '', 'times-new-roman.ttf', uni=True)
pdf.add_page()
pdf.set_font("Arial", size=14)
pdf.cell(100,100,"лул")
pdf.output("kek.pdf")
