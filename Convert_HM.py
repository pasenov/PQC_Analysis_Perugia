import tkinter
from tkinter import *
from tkinter import filedialog,StringVar
from tkinter.ttk import Frame, Button, Style
from tkinter import Tk
from tkinter.filedialog import askdirectory
import tkinter.font as font
import xlsxwriter
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

batchString = input("Batch number: ")
#For example: 34357
typeString = input("Type (2-S, PSS or PSP): ")
#For example: 2-S

for j in range(2):
    if j == 0:
        sideString = 'E'
    elif j == 1:
        sideString = 'W'
    else:
        sideString = 'Changes and progress very rarely are gifts from above. They come out of struggles from below.'

    for i in range (101):
        if i == 0:
            sampleString = '000'
        elif i == 1:
            sampleString = '001'
        elif i == 2:
            sampleString = '002'
        elif i == 3:
            sampleString = '003'
        elif i == 4:
            sampleString = '004'
        elif i == 5:
            sampleString = '005'
        elif i == 6:
            sampleString = '006'
        elif i == 7:
            sampleString = '007'
        elif i == 8:
            sampleString = '008'
        elif i == 9:
            sampleString = '009'
        elif i == 10:
            sampleString = '010'
        elif i == 11:
            sampleString = '011'
        elif i == 12:
            sampleString = '012'
        elif i == 13:
            sampleString = '013'
        elif i == 14:
            sampleString = '014'
        elif i == 15:
            sampleString = '015'
        elif i == 16:
            sampleString = '016'
        elif i == 17:
            sampleString = '017'
        elif i == 18:
            sampleString = '018'
        elif i == 19:
            sampleString = '019'
        elif i == 20:
            sampleString = '020'
        elif i == 21:
            sampleString = '021'
        elif i == 22:
            sampleString = '022'
        elif i == 23:
            sampleString = '023'
        elif i == 24:
            sampleString = '024'
        elif i == 25:
            sampleString = '025'
        elif i == 26:
            sampleString = '026'
        elif i == 27:
            sampleString = '027'
        elif i == 28:
            sampleString = '028'
        elif i == 29:
            sampleString = '029'
        elif i == 30:
            sampleString = '030'
        elif i == 31:
            sampleString = '031'
        elif i == 32:
            sampleString = '032'
        elif i == 33:
            sampleString = '033'
        elif i == 34:
            sampleString = '034'
        elif i == 35:
            sampleString = '035'
        elif i == 36:
            sampleString = '036'
        elif i == 37:
            sampleString = '037'
        elif i == 38:
            sampleString = '038'
        elif i == 39:
            sampleString = '039'
        elif i == 40:
            sampleString = '040'
        elif i == 41:
            sampleString = '041'
        elif i == 42:
            sampleString = '042'
        elif i == 43:
            sampleString = '043'
        elif i == 44:
            sampleString = '044'
        elif i == 45:
            sampleString = '045'
        elif i == 46:
            sampleString = '046'
        elif i == 47:
            sampleString = '047'
        elif i == 48:
            sampleString = '048'
        elif i == 49:
            sampleString = '049'
        elif i == 50:
            sampleString = '050'
        elif i == 51:
            sampleString = '051'
        elif i == 52:
            sampleString = '052'
        elif i == 53:
            sampleString = '053'
        elif i == 54:
            sampleString = '054'
        elif i == 55:
            sampleString = '055'
        elif i == 56:
            sampleString = '056'
        elif i == 57:
            sampleString = '057'
        elif i == 58:
            sampleString = '058'
        elif i == 59:
            sampleString = '059'
        elif i == 60:
            sampleString = '060'
        elif i == 61:
            sampleString = '061'
        elif i == 62:
            sampleString = '062'
        elif i == 63:
            sampleString = '063'
        elif i == 64:
            sampleString = '064'
        elif i == 65:
            sampleString = '065'
        elif i == 66:
            sampleString = '066'
        elif i == 67:
            sampleString = '067'
        elif i == 68:
            sampleString = '068'
        elif i == 69:
            sampleString = '069'
        elif i == 70:
            sampleString = '070'
        elif i == 71:
            sampleString = '071'
        elif i == 72:
            sampleString = '072'
        elif i == 73:
            sampleString = '073'
        elif i == 74:
            sampleString = '074'
        elif i == 75:
            sampleString = '075'
        elif i == 76:
            sampleString = '076'
        elif i == 77:
            sampleString = '077'
        elif i == 78:
            sampleString = '078'
        elif i == 79:
            sampleString = '079'
        elif i == 80:
            sampleString = '080'
        elif i == 81:
            sampleString = '081'
        elif i == 82:
            sampleString = '082'
        elif i == 83:
            sampleString = '083'
        elif i == 84:
            sampleString = '084'
        elif i == 85:
            sampleString = '085'
        elif i == 86:
            sampleString = '086'
        elif i == 87:
            sampleString = '087'
        elif i == 88:
            sampleString = '088'
        elif i == 89:
            sampleString = '089'
        elif i == 90:
            sampleString = '090'
        elif i == 91:
            sampleString = '091'
        elif i == 92:
            sampleString = '092'
        elif i == 93:
            sampleString = '093'
        elif i == 94:
            sampleString = '094'
        elif i == 95:
            sampleString = '095'
        elif i == 96:
            sampleString = '096'
        elif i == 97:
            sampleString = '097'
        elif i == 98:
            sampleString = '098'
        elif i == 99:
            sampleString = '099'
        elif i == 100:
            sampleString = '100'
        else:
            sampleString = 'You asked me once, what was in Room 101. I told you that you knew the answer already. Everyone knows it. The thing that is in Room 101 is the worst thing in the world.'

        oldPathString = 'Data/' + batchString + '/' + batchString + '_' + sampleString + '_' + typeString + '_HM_' + sideString
        newPathString = 'ConvertedData/' + batchString + '/' + batchString + '_' + sampleString + '_' + typeString + '_HM_' + sideString

        oldPath = Path(oldPathString)

        if path.exists(oldPath):
            newPath = Path(newPathString)
            pathlib.Path(newPath).mkdir(parents=True, exist_ok=True) 

            for file in os.listdir(oldPath):
                fileIn = os.fsdecode(file)
                if fileIn.endswith(".txt"):

                    #Capacitor
                    keywordsCapacitor = ['flute1_L_Capacitor', 'flute1_R_Capacitor']
                    for keyword in keywordsCapacitor:
                        if keyword in fileIn:
                            fileOut = 'PQC1-CAP_E-CV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[1:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]
                            col6 = data2[:,5]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000000)
                            colnew3a = np.multiply(col3, 1000000)
                            colnew3 = np.reciprocal(colnew3a)
                            colnew4 = col4
                            colnew5 = col5
                            colnew6 = col6

                            format1 = '%.2E', '%.2f', '%.5f', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5, colnew6])
                            #print(arr)
                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CAPCTNC_PFRD,RESSTNC_MOHM,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #FET
                    keywordsFET = ['flute1_L_FET', 'flute1_R_FET']
                    for keyword in keywordsFET:
                        if keyword in fileIn:
                            fileOut = 'PQC1-FET-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()

                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)

                            format1 = '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2])
                            #print(arr)

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP", comments="", fmt = format1)
                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #MOS
                    keywordsMOS = ['flute1_L_MOS', 'flute1_R_MOS']
                    for keyword in keywordsMOS:
                        if keyword in fileIn:
                            fileOut = 'PQC1-MOS_QUARTER-CV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[1:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]
                            col6 = data2[:,5]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000000)
                            colnew3 = np.multiply(col3, 0.000001)
                            colnew4 = col4
                            colnew5 = col5
                            colnew6 = col6

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5, colnew6])
                            #print(arr)

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CAPCTNC_PFRD,RESSTNC_MOHM,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)
                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #n+
                    keywordsnPlus = ['flute1_L_n+', 'flute1_R_n+']
                    for keyword in keywordsnPlus:
                        if keyword in fileIn:
                            fileOut = 'PQC1-VDP_STRIP-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])
                            #print(arr)

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Poly
                    keywordsPoly = ['flute1_L_Poly', 'flute1_R_Poly']
                    for keyword in keywordsPoly:
                        if keyword in fileIn:
                            fileOut = 'PQC1-VDP_POLY-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()

                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])
                            #print(arr)

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #p-stop
                    keywordspstop = ['flute1_L_pstop', 'flute1_R_pstop']
                    for keyword in keywordspstop:
                        if keyword in fileIn:
                            fileOut = 'PQC1-VDP_PSTOP-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])
                            #print(arr)

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)
                            
                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Dielectric breakdown
                    keywordsDielectric = ['flute2_L_Dielectric', 'flute2_R_Dielectric']
                    for keyword in keywordsDielectric:
                        if keyword in fileIn:
                            fileOut = 'PQC2-DIEL_NE-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[1:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])
                            #print(arr)

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #GCD
                    keywordsGCD = ['flute2_L_GCD', 'flute2_R_GCD']
                    for keyword in keywordsGCD:
                        if keyword in fileIn:
                            fileOut = 'PQC2-GCD-IGV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()

                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)

                            format1 = '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2])
                            #print(arr)

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP", comments="", fmt = format1)
                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #n+ linewidth
                    keywordsnPlusLinewidth = ['flute2_L_n+_linewidth', 'flute2_R_n+_linewidth']
                    for keyword in keywordsnPlusLinewidth:
                        if keyword in fileIn:
                            fileOut = 'PQC2-LINEWIDTH_STRIP-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Poly meander
                    keywordsPolyMeander = ['flute2_L_PolyMeander', 'flute2_R_PolyMeander']
                    for keyword in keywordsPolyMeander:
                        if keyword in fileIn:
                            fileOut = 'PQC2-R_POLY-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()

                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #p-stop linewidth
                    keywordspstopLinewidth = ['flute2_L_pstopLinewidth', 'flute2_R_pstopLinewidth']
                    for keyword in keywordspstopLinewidth:
                        if keyword in fileIn:
                            fileOut = 'PQC2-LINEWIDTH-PSTOP-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()

                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Bulk cross
                    keywordsBulkCross = ['flute3_L_BulckCross', 'flute3_R_BulckCross']
                    for keyword in keywordsBulkCross:
                        if keyword in fileIn:
                            fileOut = 'PQC3-VDP_BULK-IV.txt'

                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()

                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Diode CV
                    keywordsDiodeCV = ['flute3_L_DiodeCV', 'flute3_R_DiodeCV']
                    for keyword in keywordsDiodeCV:
                        if keyword in fileIn:
                            fileOut = 'PQC3-DIODE_HALF-CV.txt'
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[1:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]
                            col6 = data2[:,5]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000000)
                            colnew3 = np.multiply(col3, 0.000001)
                            colnew4 = col4
                            colnew5 = col5
                            colnew6 = col6

                            format1 = '%.2E', '%.2f', '%.2E', '%.2E', '%.2E', '%.2E'

                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5, colnew6])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CAPCTNC_PFRD,RESSTNC_MOHM,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Diode IV
                    keywordsDiodeIV = ['flute3_L_DiodeIV', 'flute3_R_DiodeIV']
                    for keyword in keywordsDiodeIV:
                        if keyword in fileIn:
                            fileOut = 'PQC3-DIODE_HALF-IV.txt'
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)

                            format1 = '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Metal clover
                    keywordsMetalClover = ['flute3_L_MetalCover', 'flute3_R_MetalCover']
                    for keyword in keywordsMetalClover:
                        if keyword in fileIn:
                            fileOut = 'PQC3-CLOVER_METAL-IV.txt'
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #p+ bridge
                    keywordspPlusBridge = ['flute3_L_p+Bridge', 'flute3_R_p+Bridge']
                    for keyword in keywordspPlusBridge:
                        if keyword in fileIn:
                            fileOut = 'PQC3-VDP_EDGE-IV.txt'
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #p+ cross
                    keywordspPlusCross = ['flute3_L_p+Cross', 'flute3_R_p+Cross']
                    for keyword in keywordspPlusCross:
                        if keyword in fileIn:
                            fileOut = 'PQC3-VDP_BULK-IV.txt'
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Metal meander chain
                    keywordsMetalMeanderChain = ['L_flute3_Metal_Meander_Chain', 'R_flute3_Metal_Meander_Chain']
                    for keyword in keywordsMetalMeanderChain:
                        if keyword in fileIn:
                            fileOut = 'PQC3-MEANDER_METAL-IV.txt'
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue


                    #GCD 05
                    keywordsGCD05 = ['flute4_L_GCD', 'flute4_R_GCD']
                    for keyword in keywordsGCD05:
                        if keyword in fileIn:
                            fileOut = 'PQC4-GCD05-IGV.txt'
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)

                            format1 = '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #n+ CBKR
                    keywordsnPlusCBKR = ['flute4_L_n+CBKR', 'flute4_R_n+CBKR']
                    for keyword in keywordsnPlusCBKR:
                        if keyword in fileIn:
                            fileOut = 'PQC4-CBKR_STRIP-IV.txt'
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue


                    #Poly CBKR
                    keywordspolyCBKR = ['flute4_L_polyCBKR', 'flute4_R_polyCBKR']
                    for keyword in keywordspolyCBKR:
                        if keyword in fileIn:
                            fileOut = 'PQC4-CBKR_POLY-IV.txt'
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #n+ chain
                    keywordsnPlusChain = ['L_flute4_n+_Chain', 'R_flute4_n+_Chain']
                    for keyword in keywordsnPlusChain:
                        if keyword in fileIn:
                            fileOut = 'PQC4-CC_STRIP-IV.txt' 
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #p+ chain
                    keywordspPlusChain = ['L_flute4_p+_Chain', 'R_flute4_p+_Chain']
                    for keyword in keywordspPlusChain:
                        if keyword in fileIn:
                            fileOut = 'PQC4-CC_EDGE-IV.txt'  
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue

                    #Poly chain
                    keywordsPolyChain = ['L_flute4_Poly_Chain', 'R_flute4_Poly_Chain']
                    for keyword in keywordsPolyChain:
                        if keyword in fileIn:
                            fileOut = 'PQC4-CC_POLY-IV.txt'   
                            
                            with open(oldPath/fileIn, 'r') as fin1:
                                data1 = fin1.read().splitlines(True)
                            fin1.close()
                            with open(newPath/fileIn, 'w') as fout1:
                                fout1.writelines(data1[2:])
                            fout1.close()

                            data2 = np.genfromtxt(newPath/fileIn, unpack=True).T

                            col1 = data2[:,0]
                            col2 = data2[:,1]
                            col3 = data2[:,2]
                            col4 = data2[:,3]
                            col5 = data2[:,4]

                            colnew1 = col1
                            colnew2 = np.multiply(col2, 1000000000)
                            colnew3 = col3
                            colnew4 = col4
                            colnew5 = col5

                            format1 = '%.2E', '%.2E', '%.2E', '%.2E', '%.2E'
                            arr = np.array([colnew1, colnew2, colnew3, colnew4, colnew5])

                            np.savetxt(newPath/fileOut, arr.T, delimiter=',', newline='\n', header="VOLTS,CURRNT_NAMP,TEMP_DEGC,AIR_TEMP_DEGC,RH_PRCNT", comments="", fmt = format1)

                            os.remove(newPath/fileIn)
                            continue
                        else:
                            continue
