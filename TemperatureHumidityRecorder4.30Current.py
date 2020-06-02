
################################################################
##                                                            ##
## Tempertaure and Humidity Logging in Excel                  ##
## Author: Adam Pontoni                                       ##
## Date:  May 2020                                            ##
##                                                            ##
##  Recommended to run this script via a .bat file            ##
##  through Windows Task Scheduler, set to run at midnight    ##
##                                                            ##

## Module Imports ##
import time
from serial import Serial
import platform
import pandas as pd
import xlsxwriter
import sys
import datetime
from pynput import keyboard

## Interfaces with Sensor ##
from rhusb.sensor import RHUSB

## Sets variable for calling dates ##
currentDT = datetime.datetime.now()

## Slow Printing Formula ##
def print_slow(str):
    for letter in str:
        sys.stdout.write(letter)
        sys.stdout.flush()
        time.sleep(0.2)

## Formula for Finding Average List Values ##
def averageLst(lst):
	return sum(lst) / len(lst)

		
## Delays Count by 60 Seconds ##
delay = 1
## Number of Counts (there are 1440 minutes in a day, ##
## but with the system freezing there is time "lost") ##
count = 10

## Empty Lists ##
dataTemp = []
dataHum = []
timeStamps = []
rowNumber = []

if __name__ == '__main__':
    print("Platform: {0}".format(platform.system()))
    if platform.system() == "Windows":
        device = "COM3"
    else:
        device = "/dev/ttyUSB0"
    print("Device: {0}".format(device))
    
    try:
        sens = RHUSB(device=device)
		
		## Introduction Statements ##
        print_slow("\n \nThe Program is Starting      " )
        print("""
		\n Instructions: 
		\n  Use ctrl+c to stop program. 
		\n  Program starts at Midnight and stops at 11:55PM each day. 
		\n  If computer loses power, restart it and let the program restart the next day.
		""")
		
        print("\nStarting {0} periodic readings every {1} seconds".format(count, delay))

		## Appends Temperatures and Humidities into Two Separate Lists ##
        while count:
            dataTemp.append(str('{0}'.format(sens.F())))
            dataHum.append(str('{0}'.format(sens.H())))
            timeStamps.append (time.localtime())
			
			## Provides Visuals for Program Running ##
            print("\n  Datapoint Added") 
            print("   [{0}]".format(sens.F()))
            print("   [{0}]".format(sens.H()))
			
			## Counts Down for Set Values ##
            count -= 1
            time.sleep(delay)
			
	## Throws Error Message if Disconnected Device ##
    except serial.serialutil.SerialException:
        print("Error: Unable to open RH-USB Serial device {0}.".format(device))

## Formatting Temperatures into Floats ##
excelTemp1 = [i.replace('b',"") for i in dataTemp]
excelTemp2 = [i.strip("\'") for i in excelTemp1]
excelTemp3 = [i.strip(" F") for i in excelTemp2]
excelTemp4 = [float(i) for i in excelTemp3]

## Formatting Humidities into Floats ##
excelHum1 = [i.replace('b',"") for i in dataHum]
excelHum2 = [i.replace(' %RH', "") for i in excelHum1]
excelHum3 = [i.strip("\'") for i in excelHum2]  
excelHum4 = [float(i) for i in excelHum3]

## Formats Timestamps (Change to 3 for Hours, 4 for Minutes, 5 for Seconds)
timeStamps2 = [(i[3]) for i in timeStamps]
timeStamps3 = [(i[4]) for i in timeStamps]

## Defines Formula That Performs Actual Work ##
def tempHumFormula():

	## Names Workbook Based on Month / Day / Year ##
	title = str(currentDT.month) + "." + str(currentDT.day) + "." + str(currentDT.year)

	## Opens Workbook and Applies Name ##
	workBook = xlsxwriter.Workbook("//pa-mfg.pa.local/freedisk/Adam P/tempAndHumidityArchive"
	
	+ "/" + str(currentDT.year) + "/"
	
	## Saves the Workbook to a File with Month / Year ## 
	
	## Note: The save file must already exist, be created manually, or scheduled in task scheduler ##
	 + str(currentDT.month) + "." + str(currentDT.year) + "/" + str(title) + '.xlsx')
	
	## Names the 3 Worksheets that will be in the Excel File ## 
	workSheet = workBook.add_worksheet('Raw Data')
	workSheet2 = workBook.add_worksheet('Daily Average')
	workSheet3 = workBook.add_worksheet('Daily Chart')
	
	## Row Variables ##
	rowTemp = 1
	rowHum = 1
	rowHour = 1
	rowMinute = 1

	## Column Variables ##
	columnTemp = 0
	columnHum = 1
	columnHour = 2
	columnMinute = 3
	
	## Sets Column Widths ##
	workSheet.set_column(0, 2, 20)
	workSheet.set_column(2, 4, 4)
	workSheet2.set_column(0, 5, 28)
	
	## Names Columns in Excel File ##
	## workSheet ##
	workSheet.write(0, 0, "Temperature F")
	workSheet.write(0, 1, "Relative Humidity %")
	workSheet.write(0, 2, "Hour")
	workSheet.write(0, 3, "Minute")
	## workSheet2 ##
	workSheet2.write(0, 0, "Minimum Temperature F")
	workSheet2.write(0, 1, "Average Temperature F")
	workSheet2.write(0, 2, "Maximum Temperature F")
	workSheet2.write(0, 3, "Minimum Relative Humidity %")
	workSheet2.write(0, 4, "Average Relative Humidity %")
	workSheet2.write(0, 5, "Maximum Relative Humidity %")
	
	workSheet2.write(1, 1, averageLst(excelTemp4) )
	workSheet2.write(1, 0, min(excelTemp4))
	workSheet2.write(1, 2, max(excelTemp4))
	workSheet2.write(1, 4, averageLst(excelHum4) )
	workSheet2.write(1, 3, min(excelHum4))
	workSheet2.write(1, 5, max(excelHum4))
	
	## Writes Data to Excel File ##
	for item in excelTemp4 :
	  workSheet.write(rowTemp, columnTemp, item)
	  rowTemp += 1
	  
	for item in excelHum4 :
	  workSheet.write(rowHum, columnHum, item)
	  rowHum += 1
  
	for item in timeStamps2 :
	  workSheet.write(rowHour, columnHour, item)
	  rowHour += 1
	  
	for item in timeStamps3 :
	  workSheet.write(rowMinute, columnMinute, item)
	  rowMinute += 1
	
	## Adds Chart to workSheet3 ##
	chartHeadings = ["Time", "Temperature", "Humidity"]
	chartData = [timeStamps2, excelTemp4, excelHum4]
	chart = workBook.add_chart({'type' : 'line'})
	
	## Note: format data in first_row, first_column, last_row, last_column ##
	chart.add_series({'name' : "Temperature" , 
	'values' : ['Raw Data', 1, 0,  len(excelTemp4), 0]})
	chart.add_series(
	{'name' : 'Humidity',
	'values' : ['Raw Data', 1, 1, len(excelHum4), 1] })
	
	## Sets Chart Style ##
	chart.set_title({'name' : 'Daily Temp and Humidity'})  
	chart.set_x_axis({'name' : 'Time'})
	chart.set_y_axis({'name' : 'Temp and Humidity'})
	chart.set_style(10)
	chart.set_size({'width': 1200, 'height': 400})
	workSheet3.insert_chart('A1', chart, {'x_offset' : 10, 'y_offset' :10})
	
	
	
	## Prints Completed Day Number ##  
	print_slow("\n \n \nDay " + str(currentDT.day) + " Done . . . \n")

	## Closes Workbook for Day ##  
	workBook.close() 

try:
	tempHumFormula()
		

## Sets ctrl+c to Raise Interrupt and Stop Program ##	  
except KeyboardInterrupt:
    pass


##	                                    				      ##
##	                                    				      ##
##		                                				      ## 
##	                                    				      ##
################################################################

 
