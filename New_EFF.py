import pyvisa
import pyfirmata
import numpy
import xlwt
import time
from xlwt import Workbook

import sys
sys.path.append(r'C:\Users\FZuo\OneDrive - Analog Devices, Inc\Projects\LTC3888\PYTHON')

import ADI_Instruments

board = pyfirmata.Arduino('COM5')


#Creating excel notebook
# Workbook is created
wb = Workbook()
# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(2, 2, 'Vin')
sheet1.write(2, 3, 'Iin')
sheet1.write(2, 4, 'Pin')
sheet1.write(2, 5, 'Vout')
sheet1.write(2, 6, 'Iout')
sheet1.write(2, 7, 'Pout')
sheet1.write(2, 8, 'Eff')

board.digital[2].write(1)
board.digital[3].write(1)
board.digital[4].write(1)
board.digital[5].write(1)
board.digital[6].write(1)
board.digital[7].write(1)
board.digital[8].write(1)


#Initiating VISA resources
#DMMx2, ELOADx1, PWRx1
rm = pyvisa.ResourceManager()
print(rm.list_resources())
chroma = rm.open_resource('USB0::0x1698::0x0837::008000000846::INSTR')
eload = rm.open_resource('USB0::0x0B3E::0x1004::SJ000785::INSTR')
vinDMM = rm.open_resource('USB0::0x0957::0x0607::MY53003503::INSTR')
voutDMM = rm.open_resource('USB0::0x0957::0x0607::MY45000620::INSTR')
#IinDMM = rm.open_resource('USB0::0x0957::0x0607::MY47005119::INSTR')

#Creating an array of current ranges from 320A down to 0A, increment of 1A/second
loadSteps = range(160, -1, -1)
rowCount = 164
    
#Turn on stuffs
chroma.write('SOURce:VOLTage 12.0')
chroma.write('SOUR:CURR 100.0')
chroma.write('CONFigure:OUTPut ON')
eload.write('INPut:STATe ON')
time.sleep(2)
#Step up to full load
eload.write('SOURce:CURRent 240')

time.sleep(5)


for steps in loadSteps:
    time.sleep(1)
    if steps <= 20:
        board.digital[2].write(0)
        board.digital[3].write(0)
        board.digital[4].write(0)
        board.digital[5].write(0)
        board.digital[6].write(0)
        board.digital[7].write(0)
        board.digital[8].write(0)
    if steps > 20 and steps <= 40:
        board.digital[2].write(1)
        board.digital[3].write(0)
        board.digital[4].write(0)
        board.digital[5].write(0)
        board.digital[6].write(0)
        board.digital[7].write(0)
        board.digital[8].write(0)
    if steps > 40 and steps <= 60:
        board.digital[2].write(1)
        board.digital[3].write(1)
        board.digital[4].write(0)
        board.digital[5].write(0)
        board.digital[6].write(0)
        board.digital[7].write(0)
        board.digital[8].write(0)
    if steps > 60 and steps <= 80:
        board.digital[2].write(1)
        board.digital[3].write(1)
        board.digital[4].write(1)
        board.digital[5].write(0)
        board.digital[6].write(0)
        board.digital[7].write(0)
        board.digital[8].write(0)
    if steps > 80 and steps <= 100:
        board.digital[2].write(1)
        board.digital[3].write(1)
        board.digital[4].write(1)
        board.digital[5].write(1)
        board.digital[6].write(0)
        board.digital[7].write(0)
        board.digital[8].write(0)
    if steps > 100 and steps <= 120:
        board.digital[2].write(1)
        board.digital[3].write(1)
        board.digital[4].write(1)
        board.digital[5].write(1)
        board.digital[6].write(1)
        board.digital[7].write(0)
        board.digital[8].write(0)
    if steps > 120 and steps <= 140:
        board.digital[2].write(1)
        board.digital[3].write(1)
        board.digital[4].write(1)
        board.digital[5].write(1)
        board.digital[6].write(1)
        board.digital[7].write(1)
        board.digital[8].write(0)
    if steps > 140:
        board.digital[2].write(1)
        board.digital[3].write(1)
        board.digital[4].write(1)
        board.digital[5].write(1)
        board.digital[6].write(1)
        board.digital[7].write(1)
        board.digital[8].write(1)
        
        
        
    EloadCommand = 'SOURce:CURRent ' + str(steps)
    eload.write(EloadCommand)
    time.sleep(1)
    rowCount-=1
    Iin = chroma.query_ascii_values('MEASure:CURRent?')
    Iout = eload.query_ascii_values('MEASure:CURRent?')
    Vin = vinDMM.query_ascii_values('MEASure:VOLTage?')
    Vout = voutDMM.query_ascii_values('MEASure:VOLTage?')
    #ResistorV = IinDMM.query_ascii_values('MEASure:VOLTage?')
    
    #if steps % 20 == 0
    #     numPhases = steps/20
    
    Pout = Iout[0] * Vout[0]
    Pin = Iin[0] * Vin[0]
    Eff = Pout/Pin
    print(Iin, Iout, Vin, Vout)
    
    sheet1.write(rowCount, 2, Vin[0])
    sheet1.write(rowCount, 3, Iin[0])
    sheet1.write(rowCount, 4, Pin)
    sheet1.write(rowCount, 5, Vout[0])
    sheet1.write(rowCount, 6, Iout[0])
    sheet1.write(rowCount, 7, Pout)
    sheet1.write(rowCount, 8, Eff)
    #sheet1.write(rowCount, 9, ResistorV[0])
 
#Turn stuffs off 
chroma.write('CONFigure:OUTPut OFF')
eload.write('INPut:STATe OFF')

wb.save('20A_NEW_344.xls')

