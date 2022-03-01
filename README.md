# PytonEfficiencyCollectorTool
Python script used to automatically collect efficiency data using off-the-shelf laboratory equipment through VISA commands

Software Requirements: 

Python Installation Packages: 
- python 3
- pyVISA
- numpy

National Instruments VISA download: 
https://www.ni.com/en-us/support/downloads/drivers/download.ni-visa.html#442805

After installing the NI instruments VISA software, restart the computer after being prompted to do so.

Hardware Instrumetns:
- Multimeter x 2
- power supply
- eload

Discovering your device's USB addresses: 
Connect computer to instrument via USB cable. 
Open command prompt
Type: 
python3
import pyvisa
rm = pyvisa.ResourceManager()
rm.list_resources()
.... the list of connected instruments via VISA will be listed
The USB address of the instruments will be shown in the command prompt, find the one that you want to communicate with, and paste that address in the code where you will be estbalishing a connection with the instrument.  In the code example, it will be between line 42 to line 45, paste the address of your instrument there. 


