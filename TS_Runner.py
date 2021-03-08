import openpyxl
import datetime
import pandas as pd
import configparser

from Client_TestReportMapping import *
from Emp_Syne_ClientMapping import *
from Syne_TestReportMapping import *

def getConfig(path):
    config = configparser.ConfigParser()
    config.read(path)
    return config
def getWeekDay(SyneTimesheetDate):
    monthToNum={'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}
    weekdayToName ={0:'Mon',1:'Tue',2:'Wed',3:'Thu',4:'Fri',5:'Sat',6:'Sun'}
    dateSplit = SyneTimesheetDate.split("-")
    int_day=int(dateSplit[0])
    monthSub = dateSplit[1][0:3];
    int_month = monthToNum[monthSub]
    int_Year= int(dateSplit[2])
    weekday = datetime.date(day=int_day, month=int_month, year=int_Year).weekday()
    return weekdayToName[weekday]

def TSRunner():
    inputExcel1 =getConfig(path)['Excel']['Syne']
    inputExcel2 =getConfig(path)['Excel']['Client']
    inputFormat = getConfig(path)['Excel']['InputFormat']
    outputExcel =getConfig(path)['Excel']['Output']


    dfSyneExcel = pd.read_excel(inputExcel1,"new sheet")
    dfClientExcel = pd.read_csv(inputExcel2)
    TimesheetDetail = dfSyneExcel.columns[0]
    columncount = len(dfSyneExcel.columns)
    TotalDays = columncount-5
    headerRow= ['','EMP ID','RESOURCE','CLIENT NAME','PROJECT','TASK','TOTAL']
    weakdayRow = ['', '', '', '', '', '','']
    #create header Row
    for day in range(1,TotalDays+1):
        headerRow.append(str(day))
    #create Weekday Row
    for i in range(5,TotalDays+5):
        weakdayRow.append(getWeekDay(dfSyneExcel.iat[2, i]))
    #get all unique employee Id's
    print(dfSyneExcel[TimesheetDetail].unique())
    print(dfSyneExcel[TimesheetDetail].count())

    outputData = [['', TimesheetDetail], [''], [''], headerRow, weakdayRow]
    for empId in EmpId_Name_Mapping(inputExcel1):
        l1 = []
        empIdNameProjectMapping=EmpId_Name_Project_Mapping(inputExcel1, str(empId))
        EmpId_TaskHour = Resource_Task_TotalHour_mapping(inputExcel1,inputFormat, str(empId))
        Emp_ClientName= Emp_Syne_Client_Mapping(inputFormat, empId)[empId]['CLIENT USER NAME']
        noOfSyneTask = EmpId_TaskHour[str(empId)].keys()
        noOfClientTask = get_AllClientTasks(inputExcel2,inputFormat,Emp_ClientName)
        getAllDates= get_AllDates_FromSyne(inputExcel1)
        #count=0
        counter = 0
        for task_key in noOfSyneTask:
            l2 = []
            if counter == 0:
                l2.append('Syne')
                l2.append(empId)
                l2.append(empIdNameProjectMapping[str(empId)]['ResourceName'])
                l2.append(Emp_ClientName)
                l2.append(empIdNameProjectMapping[str(empId)]['Project'])
                counter = counter + 1
            else:
                l2.append('')
                l2.append('')
                l2.append('')
                l2.append('')
                l2.append('')

            l2.append(task_key)
            l2.append(EmpId_TaskHour[str(empId)][task_key])
            for day in getAllDates:
                GetDate_Hour_Mapping = Syne_Date_Hours_Mapping(inputExcel1,inputFormat, str(empId), day)
                l2.append(GetDate_Hour_Mapping[day][task_key])
            outputData.append(l2)
        counter = 0
        for task_key in noOfClientTask:
            l3 = []
            if counter == 0:
                l3.append('Client')
                l3.append(empId)
                l3.append(empIdNameProjectMapping[str(empId)]['ResourceName'])
                l3.append(Emp_ClientName)
                l3.append(empIdNameProjectMapping[str(empId)]['Project'])
                counter = counter + 1
            else:
                l3.append('')
                l3.append('')
                l3.append('')
                l3.append('')
                l3.append('')

            l3.append(task_key)
            l3.append(get_EmpName_TotalTaskHour(inputExcel2,inputFormat,Emp_ClientName)[task_key])
            for day in getAllDates:
                GetDate_Hour_Mapping = Client_EachDate_Hour_Mapping(inputExcel2, inputFormat, Emp_ClientName, day)
                l3.append(GetDate_Hour_Mapping[day][task_key])
            outputData.append(l3)
        outputData.append(l1)


    Outputdf1 = pd.DataFrame(outputData)
    # Output = Outputdf1.style.applymap(lambda x: "background-color: lightgreen" if x == 8 else '')
    Outputdf1.to_excel(outputExcel,sheet_name='Sheet_name_1', header=False,index=False)


path="C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\properties.ini"
TSRunner()


