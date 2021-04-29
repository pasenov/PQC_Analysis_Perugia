import tkinter
from tkinter import *
from tkinter import filedialog,StringVar
from tkinter.ttk import Frame, Button, Style
from tkinter import Tk
from tkinter.filedialog import askdirectory
import tkinter.font as font
import numpy as np
import xlsxwriter
import xlrd
import sys
import os
import os.path
from os import path
import csv
import pathlib
from decimal import Decimal
from pathlib import Path
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np
import statistics

batchString = input("Batch number: ")
#For example: 34357
typeString = input("Type (2-S, PSS or PSP): ")
#For example: 2-S

mainDir = os.getcwd()

fieldhms = []

for filename in os.listdir(mainDir):
    if filename.endswith(".xlsx"):
        new_string = filename.replace(".xlsx", "")
        if "summary" not in new_string:
            fieldhms.append(new_string)
fieldhms.sort()

lengthArray = len(fieldhms)

fieldSum = ['Median', 'Standard deviation', 'Median absolute deviation']

output = 'All_flutes_summary_' + batchString + '_' + typeString + '.xlsx'
os.remove(output) if os.path.exists(output) else None
workbook = xlsxwriter.Workbook(output)

row_num1 = 0
row_num2 = 0
row_num3 = 0
row_num4 = 0

fname = []

for jname,dataname in enumerate(fieldhms):
    fname.append(dataname + ".xlsx")
    
sci_format = workbook.add_format({'num_format': '###########0.00'})
vdp_metal_format = workbook.add_format({'num_format': '###########0.0000'})
int_format = workbook.add_format({'num_format': '###########0'})

worksheet = workbook.add_worksheet('Flute1')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 25) 
worksheet.set_column('B:G', 14)
fieldnames10 = ['Flute1', 'Capacitor (pF)', 'VdP Poly (Ω/sq)', 'VdP strip (Ω/sq)', 'VdP stop (Ω/sq)', 'FET V_TH (V)', 'MOS V_fb (V)']
fieldnames11 = [' ', 'L', 'L', 'L', 'L', 'L', 'L']
for i10,data10 in enumerate(fieldnames10):
    worksheet.write(0,i10,data10,tformat)
for i11,data11 in enumerate(fieldnames11):
    worksheet.write(1,i11,data11,tformat)
for i12,data12 in enumerate(fieldhms):
    worksheet.write(i12+2,0,data12)
for i13,data13 in enumerate(fieldSum):
    worksheet.write(i13+2+lengthArray,0,data13,mformat)
a1Data = []
for jname,dataname in enumerate(fname):
    xl_workbook = xlrd.open_workbook(dataname)
    sheet = xl_workbook.sheet_by_index(0)
    a1 = [sheet.cell_value(rowx=1, colx=1), sheet.cell_value(rowx=2, colx=1), sheet.cell_value(rowx=3, colx=1), sheet.cell_value(rowx=4, colx=1), sheet.cell_value(rowx=5, colx=1), -sheet.cell_value(rowx=6, colx=1)]
    a1Data.append(a1)
    for col_num1, col_data1 in enumerate(a1):
        worksheet.write(row_num1+2, col_num1+1, col_data1, sci_format)
    row_num1 = row_num1 + 1
nbHM = row_num1
for i in range(6):
    aDataCol = [item[i] for item in a1Data]
    medMSD = statistics.median(aDataCol)
    worksheet.write(nbHM+2, i+1, medMSD, sci_format)
    stdMSD = statistics.stdev(aDataCol)
    worksheet.write(nbHM+3, i+1, stdMSD, sci_format)


worksheet = workbook.add_worksheet('Flute2')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 25)
worksheet.set_column('B:B', 17)
worksheet.set_column('C:C', 21)
worksheet.set_column('D:D', 14)
worksheet.set_column('E:E', 21)
worksheet.set_column('F:G', 17)
fieldnames20 = ['Flute2', 'n+ linewidth (μm)', 'Polysilicon meander (Ω)', 'GCD s0 (cm/s)', 'p-stop linewidth (μm)', 'Dielectric Break (V)']
fieldnames21 = [' ', 'L', 'L', 'L', 'L', 'L']
for i20,data20 in enumerate(fieldnames20):
    worksheet.write(0,i20,data20,tformat)
for i21,data21 in enumerate(fieldnames21):
    worksheet.write(1,i21,data21,tformat)
for i22,data22 in enumerate(fieldhms):
    worksheet.write(i22+2,0,data22)
for i23,data23 in enumerate(fieldSum):
    worksheet.write(i23+2+lengthArray,0,data23,mformat)
a2Data = []
for jname,dataname in enumerate(fname):
    xl_workbook = xlrd.open_workbook(dataname)
    sheet = xl_workbook.sheet_by_index(0)
    a2 = [128.5*(sheet.cell_value(rowx=3, colx=1))/(sheet.cell_value(rowx=7, colx=1)), sheet.cell_value(rowx=8, colx=1), (sheet.cell_value(rowx=9, colx=1))*0.000000000001/1.6E-019/5415000000/0.00505, 128.5*(sheet.cell_value(rowx=4, colx=1))/(sheet.cell_value(rowx=10, colx=1))]
    adiel2 = sheet.cell_value(rowx=11, colx=1)
    a2mod = [128.5*(sheet.cell_value(rowx=3, colx=1))/(sheet.cell_value(rowx=7, colx=1)), sheet.cell_value(rowx=8, colx=1), (sheet.cell_value(rowx=9, colx=1))*0.000000000001/1.6E-019/5415000000/0.00505, 128.5*(sheet.cell_value(rowx=4, colx=1))/(sheet.cell_value(rowx=10, colx=1)), sheet.cell_value(rowx=11, colx=1)]
    a2Data.append(a2mod)
    for col_num2, col_data2 in enumerate(a2):
        worksheet.write(row_num2+2, col_num2+1, col_data2, sci_format)
    worksheet.write(row_num2+2, 5, adiel2)
    row_num2 = row_num2 + 1
for i in range(4):
    aDataCol = [item[i] for item in a2Data]
    medMSD = statistics.median(aDataCol)
    worksheet.write(nbHM+2, i+1, medMSD, sci_format)
    stdMSD = statistics.stdev(aDataCol)
    worksheet.write(nbHM+3, i+1, stdMSD, sci_format)
aDataCol = [item[4] for item in a2Data]
medMSD = statistics.median(aDataCol)
worksheet.write(nbHM+2, 5, medMSD)
stdMSD = statistics.stdev(aDataCol)
worksheet.write(nbHM+3, 5, stdMSD)


worksheet = workbook.add_worksheet('Flute3')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 25) 
worksheet.set_column('B:I', 18)
fieldnames30 = ['Flute3', 'Bulk VdP (Ω/sq)', 'Diode IV (A)','Diode CV Vd (V)','VdP metal (Ω/sq)','VdP edge (Ω/sq)','Linewidth edge (μm)','Metal Meander (Ω)', 'Bulk flute3 (Ω/sq)']
fieldnames31 = [' ', 'L', 'L', 'L', 'L', 'L', 'L', 'L', 'L']
for i30,data30 in enumerate(fieldnames30):
    worksheet.write(0,i30,data30,tformat)
for i31,data31 in enumerate(fieldnames31):
    worksheet.write(1,i31,data31,tformat)
for i32,data32 in enumerate(fieldhms):
    worksheet.write(i32+2,0,data32)
for i33,data33 in enumerate(fieldSum):
    worksheet.write(i33+2+lengthArray,0,data33,mformat)
a3Data = []
for jname,dataname in enumerate(fname):
    xl_workbook = xlrd.open_workbook(dataname)
    sheet = xl_workbook.sheet_by_index(0)
    abulk3 = sheet.cell_value(rowx=12, colx=1)
    aiv3 = sheet.cell_value(rowx=13, colx=1)
    acv3 = sheet.cell_value(rowx=14, colx=1)
    acvRound3 = round(acv3, 0)
    ametal3 = sheet.cell_value(rowx=15, colx=1)
    checkp = (sheet.cell_value(rowx=17, colx=1))
    if checkp == 0:
        checkn = 1
    else:
        checkn = checkp
    a3 = [sheet.cell_value(rowx=16, colx=1), 128.5*(sheet.cell_value(rowx=16, colx=1))/checkn, sheet.cell_value(rowx=18, colx=1)]
    ares3 = 0
    if acvRound3 == 240:
        ares3 = 3502
    elif acvRound3 == 241:
        ares3 = 3488
    elif acvRound3 == 242:
        ares3 = 3473
    elif acvRound3 == 243:
        ares3 = 3459
    elif acvRound3 == 244:
        ares3 = 3445
    elif acvRound3 == 245:
        ares3 = 3431
    elif acvRound3 == 246:
        ares3 = 3417
    elif acvRound3 == 247:
        ares3 = 3403
    elif acvRound3 == 248:
        ares3 = 3389
    elif acvRound3 == 249:
        ares3 = 3375
    elif acvRound3 == 250:
        ares3 = 3362
    elif acvRound3 == 251:
        ares3 = 3349
    elif acvRound3 == 252:
        ares3 = 3335
    elif acvRound3 == 253:
        ares3 = 3322
    elif acvRound3 == 254:
        ares3 = 3309
    elif acvRound3 == 255:
        ares3 = 3296
    elif acvRound3 == 256:
        ares3 = 3283
    elif acvRound3 == 257:
        ares3 = 3270
    elif acvRound3 == 258:
        ares3 = 3258
    elif acvRound3 == 259:
        ares3 = 3245
    elif acvRound3 == 260:
        ares3 = 3233
    elif acvRound3 == 261:
        ares3 = 3220
    elif acvRound3 == 262:
        ares3 = 3208
    elif acvRound3 == 263:
        ares3 = 3196
    elif acvRound3 == 264:
        ares3 = 3184
    elif acvRound3 == 265:
        ares3 = 3172
    elif acvRound3 == 266:
        ares3 = 3160
    else:
        ares3 = -1000

    a3mod = [sheet.cell_value(rowx=12, colx=1), sheet.cell_value(rowx=13, colx=1), acv3, ametal3, sheet.cell_value(rowx=16, colx=1), 128.5*(sheet.cell_value(rowx=16, colx=1))/checkn, sheet.cell_value(rowx=18, colx=1), ares3]
    a3Data.append(a3mod)
    worksheet.write(row_num3+2, 1, abulk3, sci_format)
    worksheet.write(row_num3+2, 2, aiv3)
    worksheet.write(row_num3+2, 3, acv3, int_format)
    worksheet.write(row_num3+2, 4, ametal3, vdp_metal_format)
    for col_num3, col_data3 in enumerate(a3):
        worksheet.write(row_num3+2, col_num3+5, col_data3, sci_format)
    worksheet.write(row_num3+2, 8, ares3, int_format)
    row_num3 = row_num3 + 1
nbHM = row_num3
aDataCol = [item[0] for item in a3Data]
medMSD = statistics.median(aDataCol)
worksheet.write(nbHM+2, 1, medMSD, sci_format)
stdMSD = statistics.stdev(aDataCol)
worksheet.write(nbHM+3, 1, stdMSD, sci_format)
aDataCol = [item[1] for item in a3Data]
medMSD = statistics.median(aDataCol)
worksheet.write(nbHM+2, 2, medMSD)
stdMSD = statistics.stdev(aDataCol)
worksheet.write(nbHM+3, 2, stdMSD)
aDataCol = [item[2] for item in a3Data]
medMSD = statistics.median(aDataCol)
worksheet.write(nbHM+2, 3, medMSD, int_format)
stdMSD = statistics.stdev(aDataCol)
worksheet.write(nbHM+3, 3, stdMSD, int_format)
aDataCol = [item[3] for item in a3Data]
medMSD = statistics.median(aDataCol)
worksheet.write(nbHM+2, 4, medMSD, vdp_metal_format)
stdMSD = statistics.stdev(aDataCol)
worksheet.write(nbHM+3, 4, stdMSD, vdp_metal_format)
for i in range(4, 8):
    aDataCol = [item[i] for item in a3Data]
    medMSD = statistics.median(aDataCol)
    worksheet.write(nbHM+2, i+1, medMSD, sci_format)
    stdMSD = statistics.stdev(aDataCol)
    worksheet.write(nbHM+3, i+1, stdMSD, sci_format)

worksheet = workbook.add_worksheet('Flute4')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 25) 
worksheet.set_column('B:G', 14)
fieldnames40 = ['Flute4', 'CBKR n+ (Ω)', 'GCD s0 (cm/s)','CBKR poly (Ω)','CC poly (Ω)','CC p+ (Ω)','CC n+ (Ω)']
fieldnames41 = [' ', 'L', 'L', 'L', 'L', 'L', 'L']
for i40,data40 in enumerate(fieldnames40):
    worksheet.write(0,i40,data40,tformat)
for i41,data41 in enumerate(fieldnames41):
    worksheet.write(1,i41,data41,tformat)
for i42,data42 in enumerate(fieldhms):
    worksheet.write(i42+2,0,data42)
for i43,data43 in enumerate(fieldSum):
    worksheet.write(i43+2+lengthArray,0,data43,mformat)
a4Data = []
for jname,dataname in enumerate(fname):
    xl_workbook = xlrd.open_workbook(dataname)
    sheet = xl_workbook.sheet_by_index(0)
    ancbkr4 = sheet.cell_value(rowx=19, colx=1)
    agcd4 = (sheet.cell_value(rowx=20, colx=1))*0.000000000001/1.6E-019/5415000000/0.00723
    a4 = [sheet.cell_value(rowx=21, colx=1), sheet.cell_value(rowx=22, colx=1), sheet.cell_value(rowx=23, colx=1), sheet.cell_value(rowx=24, colx=1)]
    a4mod = [sheet.cell_value(rowx=19, colx=1), (sheet.cell_value(rowx=20, colx=1))*0.000000000001/1.6E-019/5415000000/0.00723, sheet.cell_value(rowx=21, colx=1), sheet.cell_value(rowx=22, colx=1), sheet.cell_value(rowx=23, colx=1), sheet.cell_value(rowx=24, colx=1)]
    a4Data.append(a4mod)
    worksheet.write(row_num4+2, 1, ancbkr4, sci_format)
    worksheet.write(row_num4+2, 2, agcd4, sci_format)
    for col_num4, col_data4 in enumerate(a4):
        worksheet.write(row_num4+2, col_num4+3, col_data4, sci_format)
    row_num4 = row_num4 + 1
nbHM = row_num4
for i in range(6):
    aDataCol = [item[i] for item in a4Data]
    medMSD = statistics.median(aDataCol)
    worksheet.write(nbHM+2, i+1, medMSD, sci_format)
    stdMSD = statistics.stdev(aDataCol)
    worksheet.write(nbHM+3, i+1, stdMSD, sci_format)


worksheet = workbook.add_worksheet('Acceptance')

workbook.close()
