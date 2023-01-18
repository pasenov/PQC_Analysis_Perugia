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
import scipy
from scipy.stats import median_abs_deviation


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
            if "_R" not in new_string:
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
dec_format = workbook.add_format({'num_format': '###########0.0'})
vdp_metal_format = workbook.add_format({'num_format': '###########0.0000'})
int_format = workbook.add_format({'num_format': '###########0'})

worksheet = workbook.add_worksheet('Flute1')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 25) 
worksheet.set_column('B:K', 14)
fieldnames10 = ['Flute1', 'Capacitor (pF)', 'VdP Poly (Ω/sq)', 'VdP strip (Ω/sq)', 'VdP stop (Ω/sq)', 'FET V_TH (V)', 'MOS V_fb (V)', 'VdP Poly (Ω/sq)', 'VdP strip (Ω/sq)', 'VdP stop (Ω/sq)', 'FET V_TH (V)']
fieldnames11 = [' ', 'L', 'L', 'L', 'L', 'L', 'L', 'R', 'R', 'R', 'R']
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
    a1 = [sheet.cell_value(rowx=1, colx=1), sheet.cell_value(rowx=2, colx=1), sheet.cell_value(rowx=3, colx=1), sheet.cell_value(rowx=4, colx=1), sheet.cell_value(rowx=5, colx=1), -sheet.cell_value(rowx=6, colx=1), 0, 0, 0, 0]
    a1Data.append(a1)
    for col_num1, col_data1 in enumerate(a1):
        worksheet.write(row_num1+2, col_num1+1, col_data1, sci_format)
    dataname_R = dataname.replace('.xlsx', '_R.xlsx')
    if path.exists(dataname_R):
        xl_workbook_R = xlrd.open_workbook(dataname_R)
        sheet_R = xl_workbook_R.sheet_by_index(0)
        a1 = [sheet.cell_value(rowx=1, colx=1), sheet.cell_value(rowx=2, colx=1), sheet.cell_value(rowx=3, colx=1), sheet.cell_value(rowx=4, colx=1), sheet.cell_value(rowx=5, colx=1), -sheet.cell_value(rowx=6, colx=1), sheet_R.cell_value(rowx=2, colx=1), sheet_R.cell_value(rowx=3, colx=1), sheet_R.cell_value(rowx=4, colx=1), sheet_R.cell_value(rowx=5, colx=1)]
        a1Data.append(a1)
        for col_num1, col_data1 in enumerate(a1):
            worksheet.write(row_num1+2, col_num1+1, col_data1, sci_format)
    row_num1 = row_num1 + 1
nbHM = row_num1
aDataCol = [item[0] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 1, medMSD, sci_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 1, stdMSD, sci_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 1, madMSD, sci_format)
aDataCol = [item[1] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 2, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 2, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 2, madMSD, int_format)
aDataCol = [item[2] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 3, medMSD, sci_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 3, stdMSD, sci_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 3, madMSD, sci_format)
aDataCol = [item[3] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 4, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 4, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 4, madMSD, int_format)
aDataCol = [item[4] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 5, medMSD, sci_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 5, stdMSD, sci_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 5, madMSD, sci_format)
aDataCol = [item[5] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 6, medMSD, sci_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 6, stdMSD, sci_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 6, madMSD, sci_format)
aDataCol = [item[6] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 7, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 7, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 7, madMSD, int_format)
aDataCol = [item[7] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 8, medMSD, sci_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 8, stdMSD, sci_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 8, madMSD, sci_format)
aDataCol = [item[8] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 9, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 9, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 9, madMSD, int_format)
aDataCol = [item[9] for item in a1Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 10, medMSD, sci_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 10, stdMSD, sci_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 10, madMSD, sci_format)
     
worksheet = workbook.add_worksheet('Flute2')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 25)
worksheet.set_column('B:B', 17)
worksheet.set_column('C:C', 21)
worksheet.set_column('D:E', 14)
worksheet.set_column('F:F', 21)
worksheet.set_column('G:G', 17)
fieldnames20 = ['Flute2', 'n+ linewidth (μm)', 'Polysilicon meander (Ω)', 'GCD s0 (cm/s)', 'GCD V_FB (V)', 'p-stop linewidth (μm)', 'Dielectric Break (V)']
fieldnames21 = [' ', 'L', 'L', 'L', 'L', 'L', 'L']
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
    a2 = [128.5*(sheet.cell_value(rowx=3, colx=1))/(sheet.cell_value(rowx=7, colx=1)), sheet.cell_value(rowx=8, colx=1), (sheet.cell_value(rowx=9, colx=1))*0.000000000001/1.6E-019/5415000000/0.00505, 5.00 + sheet.cell_value(rowx=10, colx=1), 128.5*(sheet.cell_value(rowx=4, colx=1))/(sheet.cell_value(rowx=11, colx=1))]
    adiel2 = sheet.cell_value(rowx=12, colx=1)
    a2mod = [128.5*(sheet.cell_value(rowx=3, colx=1))/(sheet.cell_value(rowx=7, colx=1)), sheet.cell_value(rowx=8, colx=1), (sheet.cell_value(rowx=9, colx=1))*0.000000000001/1.6E-019/5415000000/0.00505, 5.00 + sheet.cell_value(rowx=10, colx=1), 128.5*(sheet.cell_value(rowx=4, colx=1))/(sheet.cell_value(rowx=11, colx=1)), sheet.cell_value(rowx=12, colx=1)]
    a2Data.append(a2mod)
    for col_num2, col_data2 in enumerate(a2):
        worksheet.write(row_num2+2, col_num2+1, col_data2, sci_format)
    worksheet.write(row_num2+2, 6, adiel2)
    row_num2 = row_num2 + 1
aDataCol = [item[0] for item in a2Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 1, medMSD, sci_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 1, stdMSD, sci_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x!= 0.0])
worksheet.write(nbHM+4, 1, madMSD, sci_format)
aDataCol = [item[1] for item in a2Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 2, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 2, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x!= 0.0])
worksheet.write(nbHM+4, 2, madMSD, int_format)
for i in range(2, 5):
    aDataCol = [item[i] for item in a2Data]
    medMSD = statistics.median(x for x in aDataCol if x != 0.0)
    worksheet.write(nbHM+2, i+1, medMSD, sci_format)
    stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
    worksheet.write(nbHM+3, i+1, stdMSD, sci_format)
    madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x!= 0.0])
    worksheet.write(nbHM+4, i+1, madMSD, sci_format)
aDataCol = [item[5] for item in a2Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 6, medMSD)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 6, stdMSD)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 6, madMSD)

worksheet = workbook.add_worksheet('Flute3')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 25) 
worksheet.set_column('B:I', 18)
fieldnames30 = ['Flute3', 'Bulk VdP (Ω.cm)', 'Diode IV (A)','Diode CV Vd (V)','VdP metal (Ω/sq)','VdP edge (Ω/sq)','Linewidth edge (μm)','Metal Meander (Ω)', 'Bulk flute3 (Ω.cm)']
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
    abulk3 = sheet.cell_value(rowx=13, colx=1)
    aiv3 = sheet.cell_value(rowx=14, colx=1)
    acv3 = sheet.cell_value(rowx=15, colx=1)
    acvRound3 = round(acv3, 0)
    ametal3 = sheet.cell_value(rowx=16, colx=1)
    checkp = (sheet.cell_value(rowx=18, colx=1))
    if checkp == 0:
        checkn = 1
    else:
        checkn = checkp
    a3 = [sheet.cell_value(rowx=17, colx=1), 128.5*(sheet.cell_value(rowx=17, colx=1))/checkn, sheet.cell_value(rowx=19, colx=1)]
    ares3 = 0
    if acvRound3 == 100:
        ares3 = 8405
    elif acvRound3 == 101:
        ares3 = 8322
    elif acvRound3 == 102:
        ares3 = 8240
    elif acvRound3 == 103:
        ares3 = 8160
    elif acvRound3 == 104:
        ares3 = 8082
    elif acvRound3 == 105:
        ares3 = 8005
    elif acvRound3 == 106:
        ares3 = 7929
    elif acvRound3 == 107:
        ares3 = 7855
    elif acvRound3 == 108:
        ares3 = 7782
    elif acvRound3 == 109:
        ares3 = 7711
    elif acvRound3 == 110:
        ares3 = 7641
    elif acvRound3 == 111:
        ares3 = 7572
    elif acvRound3 == 112:
        ares3 = 7504
    elif acvRound3 == 113:
        ares3 = 7438
    elif acvRound3 == 114:
        ares3 = 7373
    elif acvRound3 == 115:
        ares3 = 7309
    elif acvRound3 == 116:
        ares3 = 7246
    elif acvRound3 == 117:
        ares3 = 7184
    elif acvRound3 == 118:
        ares3 = 7123
    elif acvRound3 == 119:
        ares3 = 7063
    elif acvRound3 == 120:
        ares3 = 7004
    elif acvRound3 == 121:
        ares3 = 6946
    elif acvRound3 == 122:
        ares3 = 6889
    elif acvRound3 == 123:
        ares3 = 6833
    elif acvRound3 == 124:
        ares3 = 6778
    elif acvRound3 == 125:
        ares3 = 6724
    elif acvRound3 == 126:
        ares3 = 6670
    elif acvRound3 == 127:
        ares3 = 6618
    elif acvRound3 == 128:
        ares3 = 6566
    elif acvRound3 == 129:
        ares3 = 6515
    elif acvRound3 == 130:
        ares3 = 6465
    elif acvRound3 == 131:
        ares3 = 6416
    elif acvRound3 == 132:
        ares3 = 6367
    elif acvRound3 == 133:
        ares3 = 6319
    elif acvRound3 == 134:
        ares3 = 6272
    elif acvRound3 == 135:
        ares3 = 6226
    elif acvRound3 == 136:
        ares3 = 6180
    elif acvRound3 == 137:
        ares3 = 6135
    elif acvRound3 == 138:
        ares3 = 6090
    elif acvRound3 == 139:
        ares3 = 6047
    elif acvRound3 == 140:
        ares3 = 6003
    elif acvRound3 == 141:
        ares3 = 5961
    elif acvRound3 == 142:
        ares3 = 5919
    elif acvRound3 == 143:
        ares3 = 5878
    elif acvRound3 == 144:
        ares3 = 5837
    elif acvRound3 == 145:
        ares3 = 5796
    elif acvRound3 == 146:
        ares3 = 5757
    elif acvRound3 == 147:
        ares3 = 5718
    elif acvRound3 == 148:
        ares3 = 5679
    elif acvRound3 == 149:
        ares3 = 5641
    elif acvRound3 == 150:
        ares3 = 5603
    elif acvRound3 == 151:
        ares3 = 5566
    elif acvRound3 == 152:
        ares3 = 5529
    elif acvRound3 == 153:
        ares3 = 5493
    elif acvRound3 == 154:
        ares3 = 5458
    elif acvRound3 == 155:
        ares3 = 5422
    elif acvRound3 == 156:
        ares3 = 5388
    elif acvRound3 == 157:
        ares3 = 5353
    elif acvRound3 == 158:
        ares3 = 5320
    elif acvRound3 == 159:
        ares3 = 5286
    elif acvRound3 == 160:
        ares3 = 5253
    elif acvRound3 == 161:
        ares3 = 5220
    elif acvRound3 == 162:
        ares3 = 5188
    elif acvRound3 == 163:
        ares3 = 5157
    elif acvRound3 == 164:
        ares3 = 5125
    elif acvRound3 == 165:
        ares3 = 5094
    elif acvRound3 == 166:
        ares3 = 5063
    elif acvRound3 == 167:
        ares3 = 5033
    elif acvRound3 == 168:
        ares3 = 5003
    elif acvRound3 == 169:
        ares3 = 4973
    elif acvRound3 == 170:
        ares3 = 4933
    elif acvRound3 == 171:
        ares3 = 4915
    elif acvRound3 == 172:
        ares3 = 4887
    elif acvRound3 == 173:
        ares3 = 4858
    elif acvRound3 == 174:
        ares3 = 4830
    elif acvRound3 == 175:
        ares3 = 4803
    elif acvRound3 == 176:
        ares3 = 4775
    elif acvRound3 == 177:
        ares3 = 4749
    elif acvRound3 == 178:
        ares3 = 4722
    elif acvRound3 == 179:
        ares3 = 4695
    elif acvRound3 == 180:
        ares3 = 4669
    elif acvRound3 == 181:
        ares3 = 4644
    elif acvRound3 == 182:
        ares3 = 4618
    elif acvRound3 == 183:
        ares3 = 4593
    elif acvRound3 == 184:
        ares3 = 4568
    elif acvRound3 == 185:
        ares3 = 4543
    elif acvRound3 == 186:
        ares3 = 4519
    elif acvRound3 == 187:
        ares3 = 4495
    elif acvRound3 == 188:
        ares3 = 4471
    elif acvRound3 == 189:
        ares3 = 4447
    elif acvRound3 == 190:
        ares3 = 4424
    elif acvRound3 == 191:
        ares3 = 4400
    elif acvRound3 == 192:
        ares3 = 4378
    elif acvRound3 == 193:
        ares3 = 4355
    elif acvRound3 == 194:
        ares3 = 4332
    elif acvRound3 == 195:
        ares3 = 4310
    elif acvRound3 == 196:
        ares3 = 4288
    elif acvRound3 == 197:
        ares3 = 4266
    elif acvRound3 == 198:
        ares3 = 4245
    elif acvRound3 == 199:
        ares3 = 4224 
    elif acvRound3 == 200:
        ares3 = 4202
    elif acvRound3 == 201:
        ares3 = 4182
    elif acvRound3 == 202:
        ares3 = 4161
    elif acvRound3 == 203:
        ares3 = 4140
    elif acvRound3 == 204:
        ares3 = 4120
    elif acvRound3 == 205:
        ares3 = 4100
    elif acvRound3 == 206:
        ares3 = 4080
    elif acvRound3 == 207:
        ares3 = 4060
    elif acvRound3 == 208:
        ares3 = 4041
    elif acvRound3 == 209:
        ares3 = 4021
    elif acvRound3 == 210:
        ares3 = 4002
    elif acvRound3 == 211:
        ares3 = 3983
    elif acvRound3 == 212:
        ares3 = 3965
    elif acvRound3 == 213:
        ares3 = 3946
    elif acvRound3 == 214:
        ares3 = 3928
    elif acvRound3 == 215:
        ares3 = 3909
    elif acvRound3 == 216:
        ares3 = 3891
    elif acvRound3 == 217:
        ares3 = 3873
    elif acvRound3 == 218:
        ares3 = 3855
    elif acvRound3 == 219:
        ares3 = 3838
    elif acvRound3 == 220:
        ares3 = 3820
    elif acvRound3 == 221:
        ares3 = 3803
    elif acvRound3 == 222:
        ares3 = 3786
    elif acvRound3 == 223:
        ares3 = 3769
    elif acvRound3 == 224:
        ares3 = 3752
    elif acvRound3 == 225:
        ares3 = 3735
    elif acvRound3 == 226:
        ares3 = 3719
    elif acvRound3 == 227:
        ares3 = 3703
    elif acvRound3 == 228:
        ares3 = 3686
    elif acvRound3 == 229:
        ares3 = 3670
    elif acvRound3 == 230:
        ares3 = 3654
    elif acvRound3 == 231:
        ares3 = 3638
    elif acvRound3 == 232:
        ares3 = 3623
    elif acvRound3 == 233:
        ares3 = 3607
    elif acvRound3 == 234:
        ares3 = 3592
    elif acvRound3 == 235:
        ares3 = 3577
    elif acvRound3 == 236:
        ares3 = 3561
    elif acvRound3 == 237:
        ares3 = 3546
    elif acvRound3 == 238:
        ares3 = 3531
    elif acvRound3 == 239:
        ares3 = 3517
    elif acvRound3 == 240:
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
    elif acvRound3 == 267:
        ares3 = 3148
    elif acvRound3 == 268:
        ares3 = 3136
    elif acvRound3 == 269:
        ares3 = 3124
    elif acvRound3 == 270:
        ares3 = 3113
    elif acvRound3 == 271:
        ares3 = 3101
    elif acvRound3 == 272:
        ares3 = 3090
    elif acvRound3 == 273:
        ares3 = 3079
    elif acvRound3 == 274:
        ares3 = 3067
    elif acvRound3 == 275:
        ares3 = 3056
    elif acvRound3 == 276:
        ares3 = 3045
    elif acvRound3 == 277:
        ares3 = 3034
    elif acvRound3 == 278:
        ares3 = 3023
    elif acvRound3 == 279:
        ares3 = 3013
    elif acvRound3 == 280:
        ares3 = 3002
    elif acvRound3 == 281:
        ares3 = 2991
    elif acvRound3 == 282:
        ares3 = 2980
    elif acvRound3 == 283:
        ares3 = 2970
    elif acvRound3 == 284:
        ares3 = 2959
    elif acvRound3 == 285:
        ares3 = 2949
    elif acvRound3 == 286:
        ares3 = 2939
    elif acvRound3 == 287:
        ares3 = 2929
    elif acvRound3 == 288:
        ares3 = 2918
    elif acvRound3 == 289:
        ares3 = 2908
    elif acvRound3 == 290:
        ares3 = 2898
    elif acvRound3 == 291:
        ares3 = 2888
    elif acvRound3 == 292:
        ares3 = 2878
    elif acvRound3 == 293:
        ares3 = 2869
    elif acvRound3 == 294:
        ares3 = 2859
    elif acvRound3 == 295:
        ares3 = 2849
    elif acvRound3 == 296:
        ares3 = 2839
    elif acvRound3 == 297:
        ares3 = 2830
    elif acvRound3 == 298:
        ares3 = 2820
    elif acvRound3 == 299:
        ares3 = 2811
    elif acvRound3 == 300:
        ares3 = 2802
    elif acvRound3 == 301:
        ares3 = 2792
    elif acvRound3 == 302:
        ares3 = 2783
    elif acvRound3 == 303:
        ares3 = 2773
    elif acvRound3 == 304:
        ares3 = 2765
    elif acvRound3 == 305:
        ares3 = 2756
    elif acvRound3 == 306:
        ares3 = 2747
    elif acvRound3 == 307:
        ares3 = 2738
    elif acvRound3 == 308:
        ares3 = 2729
    elif acvRound3 == 309:
        ares3 = 2720
    elif acvRound3 == 310:
        ares3 = 2711
    elif acvRound3 == 311:
        ares3 = 2703
    elif acvRound3 == 312:
        ares3 = 2694
    elif acvRound3 == 313:
        ares3 = 2685
    elif acvRound3 == 314:
        ares3 = 2677
    elif acvRound3 == 315:
        ares3 = 2668
    elif acvRound3 == 316:
        ares3 = 2660
    elif acvRound3 == 317:
        ares3 = 2651
    elif acvRound3 == 318:
        ares3 = 2643
    elif acvRound3 == 319:
        ares3 = 2635
    elif acvRound3 == 320:
        ares3 = 2627
    elif acvRound3 == 321:
        ares3 = 2618
    elif acvRound3 == 322:
        ares3 = 2610
    elif acvRound3 == 323:
        ares3 = 2602
    elif acvRound3 == 324:
        ares3 = 2594
    elif acvRound3 == 325:
        ares3 = 2586
    elif acvRound3 == 326:
        ares3 = 2578
    elif acvRound3 == 327:
        ares3 = 2570
    elif acvRound3 == 328:
        ares3 = 2562
    elif acvRound3 == 329:
        ares3 = 2555
    elif acvRound3 == 330:
        ares3 = 2547
    elif acvRound3 == 331:
        ares3 = 2539
    elif acvRound3 == 332:
        ares3 = 2532
    elif acvRound3 == 333:
        ares3 = 2524
    elif acvRound3 == 334:
        ares3 = 2516
    elif acvRound3 == 335:
        ares3 = 2509
    elif acvRound3 == 336:
        ares3 = 2501
    elif acvRound3 == 337:
        ares3 = 2494
    elif acvRound3 == 338:
        ares3 = 2487
    elif acvRound3 == 339:
        ares3 = 2479
    elif acvRound3 == 340:
        ares3 = 2472
    elif acvRound3 == 341:
        ares3 = 2465
    elif acvRound3 == 342:
        ares3 = 2458
    elif acvRound3 == 343:
        ares3 = 2450
    elif acvRound3 == 344:
        ares3 = 2443
    elif acvRound3 == 345:
        ares3 = 2436
    elif acvRound3 == 346:
        ares3 = 2429
    elif acvRound3 == 347:
        ares3 = 2422
    elif acvRound3 == 348:
        ares3 = 2415
    elif acvRound3 == 349:
        ares3 = 2408
    elif acvRound3 == 350:
        ares3 = 2401
    elif acvRound3 == 351:
        ares3 = 2395
    elif acvRound3 == 352:
        ares3 = 2388
    elif acvRound3 == 353:
        ares3 = 2381
    elif acvRound3 == 354:
        ares3 = 2374
    elif acvRound3 == 355:
        ares3 = 2368
    elif acvRound3 == 356:
        ares3 = 2361
    elif acvRound3 == 357:
        ares3 = 2354
    elif acvRound3 == 358:
        ares3 = 2348
    elif acvRound3 == 359:
        ares3 = 2341
    elif acvRound3 == 360:
        ares3 = 2335
    elif acvRound3 == 361:
        ares3 = 2328
    elif acvRound3 == 362:
        ares3 = 2322
    elif acvRound3 == 363:
        ares3 = 2315
    elif acvRound3 == 364:
        ares3 = 2309
    elif acvRound3 == 365:
        ares3 = 2303
    elif acvRound3 == 366:
        ares3 = 2296
    elif acvRound3 == 367:
        ares3 = 2290
    elif acvRound3 == 368:
        ares3 = 2284
    elif acvRound3 == 369:
        ares3 = 2278
    elif acvRound3 == 370:
        ares3 = 2272
    elif acvRound3 == 371:
        ares3 = 2265
    elif acvRound3 == 372:
        ares3 = 2259
    elif acvRound3 == 373:
        ares3 = 2253
    elif acvRound3 == 374:
        ares3 = 2247
    elif acvRound3 == 375:
        ares3 = 2241
    elif acvRound3 == 376:
        ares3 = 2235
    elif acvRound3 == 377:
        ares3 = 2229
    elif acvRound3 == 378:
        ares3 = 2224
    elif acvRound3 == 379:
        ares3 = 2218
    elif acvRound3 == 380:
        ares3 = 2212
    elif acvRound3 == 381:
        ares3 = 2206
    elif acvRound3 == 382:
        ares3 = 2200
    elif acvRound3 == 383:
        ares3 = 2195
    elif acvRound3 == 384:
        ares3 = 2189
    elif acvRound3 == 385:
        ares3 = 2183
    elif acvRound3 == 386:
        ares3 = 2177
    elif acvRound3 == 387:
        ares3 = 2172
    elif acvRound3 == 388:
        ares3 = 2166
    elif acvRound3 == 389:
        ares3 = 2161
    elif acvRound3 == 390:
        ares3 = 2155
    elif acvRound3 == 391:
        ares3 = 2150
    elif acvRound3 == 392:
        ares3 = 2144
    elif acvRound3 == 393:
        ares3 = 2139
    elif acvRound3 == 394:
        ares3 = 2133
    elif acvRound3 == 395:
        ares3 = 2128
    elif acvRound3 == 396:
        ares3 = 2122
    elif acvRound3 == 397:
        ares3 = 2117
    elif acvRound3 == 398:
        ares3 = 2112
    elif acvRound3 == 399:
        ares3 = 2107
    elif acvRound3 == 400:
        ares3 = 2101
    else:
        ares3 = -1000

    a3mod = [sheet.cell_value(rowx=13, colx=1), sheet.cell_value(rowx=14, colx=1), acv3, ametal3, sheet.cell_value(rowx=17, colx=1), 128.5*(sheet.cell_value(rowx=17, colx=1))/checkn, sheet.cell_value(rowx=19, colx=1), ares3]
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
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 1, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 1, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 1, madMSD, int_format)
aDataCol = [item[1] for item in a3Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 2, medMSD)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 2, stdMSD)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 2, madMSD)
aDataCol = [item[2] for item in a3Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 3, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 3, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 3, madMSD, int_format)
aDataCol = [item[3] for item in a3Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 4, medMSD, vdp_metal_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 4, stdMSD, vdp_metal_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 4, madMSD, vdp_metal_format)
aDataCol = [item[4] for item in a3Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 5, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 5, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 5, madMSD, int_format)
aDataCol = [item[5] for item in a3Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 6, medMSD, sci_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 6, stdMSD, sci_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 6, madMSD, sci_format)
for i in range(6, 7):
    aDataCol = [item[i] for item in a3Data]
    medMSD = statistics.median(x for x in aDataCol if x != 0.0)
    worksheet.write(nbHM+2, i+1, medMSD, dec_format)
    stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
    worksheet.write(nbHM+3, i+1, stdMSD, dec_format)
    madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
    worksheet.write(nbHM+4, i+1, madMSD, dec_format)
aDataCol = [item[7] for item in a3Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+2, 8, medMSD, int_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0)
worksheet.write(nbHM+3, 8, stdMSD, int_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0])
worksheet.write(nbHM+4, 8, madMSD, int_format)
    
worksheet = workbook.add_worksheet('Flute4')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_column('A:A', 25) 
worksheet.set_column('B:H', 14)
fieldnames40 = ['Flute4', 'CBKR n+ (Ω)', 'GCD s0 (cm/s)', 'GCD V_FB (V)', 'CBKR poly (Ω)', 'CC poly (Ω)', 'CC p+ (Ω)', 'CC n+ (Ω)']
fieldnames41 = [' ', 'L', 'L', 'L', 'L', 'L', 'L', 'L']
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
    ancbkr4 = sheet.cell_value(rowx=20, colx=1)
    agcd4 = (sheet.cell_value(rowx=21, colx=1))*0.000000000001/1.6E-019/5415000000/0.00723
    a4 = [5.00 + sheet.cell_value(rowx=22, colx=1), sheet.cell_value(rowx=23, colx=1), sheet.cell_value(rowx=24, colx=1), sheet.cell_value(rowx=25, colx=1), sheet.cell_value(rowx=26, colx=1)]
    a4mod = [sheet.cell_value(rowx=20, colx=1), (sheet.cell_value(rowx=21, colx=1))*0.000000000001/1.6E-019/5415000000/0.00723, 5.00 + sheet.cell_value(rowx=22, colx=1), sheet.cell_value(rowx=23, colx=1), sheet.cell_value(rowx=24, colx=1), sheet.cell_value(rowx=25, colx=1), sheet.cell_value(rowx=26, colx=1)]
    a4Data.append(a4mod)
    worksheet.write(row_num4+2, 1, ancbkr4, sci_format)
    worksheet.write(row_num4+2, 2, agcd4, sci_format)
    for col_num4, col_data4 in enumerate(a4):
        worksheet.write(row_num4+2, col_num4+3, col_data4, sci_format)
    row_num4 = row_num4 + 1
nbHM = row_num4
aDataCol = [item[0] for item in a4Data]
medMSD = statistics.median(x for x in aDataCol if x != 0.0 and x != 5.0)
worksheet.write(nbHM+2, 1, medMSD, dec_format)
stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0 and x != 5.0)
worksheet.write(nbHM+3, 1, stdMSD, dec_format)
madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0 and x != 5.0])
worksheet.write(nbHM+4, 1, madMSD, dec_format)
for i in range(1, 3):
    aDataCol = [item[i] for item in a4Data]
    medMSD = statistics.median(x for x in aDataCol if x != 0.0 and x != 5.0)
    worksheet.write(nbHM+2, i+1, medMSD, sci_format)
    stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0 and x != 5.0)
    worksheet.write(nbHM+3, i+1, stdMSD, sci_format)
    madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0 and x != 5.0])
    worksheet.write(nbHM+4, i+1, madMSD, sci_format)
for i in range(3, 7):
    aDataCol = [item[i] for item in a4Data]
    medMSD = statistics.median(x for x in aDataCol if x != 0.0 and x != 5.0)
    worksheet.write(nbHM+2, i+1, medMSD, int_format)
    stdMSD = statistics.stdev(x for x in aDataCol if x != 0.0 and x != 5.0)
    worksheet.write(nbHM+3, i+1, stdMSD, int_format)
    madMSD = statistics.median([abs(x - medMSD) for x in aDataCol if x != 0.0 and x != 5.0])
    worksheet.write(nbHM+4, i+1, madMSD, int_format)

worksheet = workbook.add_worksheet('Acceptance')
tformat = workbook.add_format({'bold': True})
tformat.set_align('center')
mformat = workbook.add_format({'bold': True})
worksheet.set_row(0, 40)
worksheet.set_column('A:A', 14) 
worksheet.set_column('B:C', 15)
worksheet.set_column('D:D', 21)
worksheet.set_column('E:F', 16)
worksheet.set_column('G:G', 14)
worksheet.set_column('H:H', 19)
fieldnames50 = ['', 'Diode full \ndepletion voltage \nVfd < 350 V', 'Flatband voltage \n|Vfb| < 5 V', 'Thin oxide \nthickness (capacitance) \nCAC > 1.2 pF cm-2 μm-1', 'Strip implant \nRstrip < 250 Ω/sq', 'Aluminium strip \nRalu < 30 mΩ/sq', 'Bulk resistivity \n> 2.7 kΩ.cm', 'Dielectric breakdown \nVdiel > 150 V']
fieldnames51 = [' ', 'L', 'L', 'L', 'L', 'L', 'L', 'L']
for i50,data50 in enumerate(fieldnames50):
    worksheet.write(0,i50,data50,tformat)
for i51,data51 in enumerate(fieldnames51):
    worksheet.write(1,i51,data51,tformat)
for i52,data52 in enumerate(fieldhms):
    worksheet.write(i52+2,0,data52)
    for j53 in range(1, 8):
        worksheet.write(i52+2,j53,'OK')

workbook.close()
