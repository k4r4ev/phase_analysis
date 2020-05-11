import openpyxl
import os

from config import Config
from libs import *

print("Phase analysis program. Sponsored by NATO & NASA")

workbook = openpyxl.load_workbook(Config.workbook)
if Config.sheet != '':
    source_sheet = workbook[Config.sheet]
else:
    source_sheet = workbook[workbook.sheetnames[0]]

import_source_data(workbook, source_sheet, Config)

sheet = workbook[Config.output_sheet]
calculate_derivative(sheet, Config)
quasicycles = get_quasicycles(sheet, Config)

q_index = 0
for quasicycle in quasicycles:
    sheet.add_chart(create_diagram(quasicycle), str(sheet.cell(q_index * 15 + 1, 5).coordinate))
    q_index += 1

os.remove(Config.workbook)
workbook.save(Config.workbook)
