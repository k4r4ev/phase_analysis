import openpyxl
import os

from quasicycle import Quasicycle
from config import Config
from libs import *

print("Phase analysis program. Sponsored by NATO & NASA")

workbook = openpyxl.load_workbook(Config.sheet)
sheet = workbook[Config.workbook]

calculate_derivative(sheet, Config)

quasicycles = [Quasicycle(sheet, "Квазицикл1", 1, 1, 10), Quasicycle(sheet, "Квазицикл2", 10, 1, 10),
               Quasicycle(sheet, "Квазицикл3", 20, 1, 10)]

sheet.add_chart(create_diagram(quasicycles[0]), str(sheet.cell(5, 5).coordinate))
sheet.add_chart(create_diagram(quasicycles[1]), str(sheet.cell(20, 5).coordinate))
sheet.add_chart(create_diagram(quasicycles[2]), str(sheet.cell(35, 5).coordinate))

os.remove(Config.sheet)
workbook.save(Config.sheet)
