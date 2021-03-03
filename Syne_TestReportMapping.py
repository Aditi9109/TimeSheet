import openpyxl
import datetime
import pandas as pd


def EmpId_Name_Mapping(inputExcel):
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")
    empIdNameMapping = {}
    for index, row in dfSyneExcel.iterrows():
        if isinstance(row[0], int):
            empIdNameMapping[str(row[0])]= str(row[1])
    return empIdNameMapping

def EmpId_Name_Project_Mapping(inputExcel,EmpID):
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")
    empIdNameProjectMapping = {}
    for i in range(1, len(dfSyneExcel)):
        if (str(dfSyneExcel.values[i][0]) == EmpID):
            resourceProject ={}
            resourceProject['ResourceName'] = dfSyneExcel.values[i][1]
            resourceProject['Project'] = dfSyneExcel.values[i][2]
            empIdNameProjectMapping[EmpID] = resourceProject
    return empIdNameProjectMapping

def get_AllDates_FromSyne(inputExcel):
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")
    AllDates=[]
    for i in range(5, len(dfSyneExcel.columns)):
        AllDates.append(dfSyneExcel.values[2][i])
    return AllDates

def Resource_Task_TotalHour_mapping(inputExcel,inputFormat,EmpID):
    headerRow = ['Billable', 'Leave', 'Public Holiday', 'Training', 'Admin/Other']
    Syne_TaskHourMap = {}
    taskHour_dict = {}
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")
    dfFormatSyne = pd.read_excel(inputFormat, "Synechron")
    billable = [x for x in dfFormatSyne['Billable'].tolist() if isinstance(x, str)]
    leave = [x for x in dfFormatSyne['Leave'].tolist() if isinstance(x, str)]
    totalBillableHour = 0
    totalLeaveHour = 0
    total_Hour = 0
    for i in range(1, len(dfSyneExcel)):
        if(str(dfSyneExcel.values[i][0])==EmpID):
            taskName = dfSyneExcel.values[i][3]
            TotalHours = dfSyneExcel.values[i][4]
            if taskName in billable:
                Task = "Billable"
                totalBillableHour = totalBillableHour + TotalHours
                taskHour_dict[Task] = totalBillableHour
            elif taskName in leave:
                Task = "Leave"
                totalLeaveHour = totalLeaveHour + TotalHours
                taskHour_dict[Task]= totalLeaveHour
        Syne_TaskHourMap[EmpID] = taskHour_dict
    return Syne_TaskHourMap

def Syne_Date_Hours_Mapping(inputExcel,inputFormat, EmpID, Date):
    Syne_date_taskHourMap = {}
    TaskDayHour_dict = {}
    dfSyneExcel = pd.read_excel(inputExcel, "new sheet")
    dfFormatSyne = pd.read_excel(inputFormat, "Synechron")
    billable = [x for x in dfFormatSyne['Billable'].tolist() if isinstance(x, str)]
    leave = [x for x in dfFormatSyne['Leave'].tolist() if isinstance(x, str)]
    dateColNo = 0
    totalBillableHour= 0
    totalLeaveHour= 0
    for j in range(1, len(dfSyneExcel.columns)):
        if (str(dfSyneExcel.values[2][j]) == Date):
            dateColNo = j
            break
    for i in range(1, len(dfSyneExcel)):
        if (str(dfSyneExcel.values[i][0]) == EmpID):
            taskName = dfSyneExcel.values[i][3]
            hourValue = dfSyneExcel.values[i][dateColNo]
            if taskName in billable:
                Task = "Billable"
                if isinstance(hourValue, int) or isinstance(hourValue, float):
                    totalBillableHour = totalBillableHour + hourValue
                    TaskDayHour_dict[Task] = totalBillableHour
                elif isinstance(hourValue, str):
                    if 'Leave' in TaskDayHour_dict.keys():
                        if isinstance(TaskDayHour_dict[Task],int) or isinstance(TaskDayHour_dict[Task],float):
                            pass
                    else:
                        TaskDayHour_dict[Task] = hourValue
            elif taskName in leave:
                Task = "Leave"
                if isinstance(hourValue, int) or isinstance(hourValue, float):
                    totalLeaveHour = totalLeaveHour + hourValue
                    TaskDayHour_dict[Task] = totalLeaveHour
                elif isinstance(hourValue, str):
                    if 'Leave' in TaskDayHour_dict.keys():
                        if isinstance(TaskDayHour_dict[Task],int) or isinstance(TaskDayHour_dict[Task],float):
                            pass
                    else:
                        TaskDayHour_dict[Task] = hourValue

            #hourValue = dfSyneExcel.values[i][dateColNo]
            #TaskDayHour_dict[Task] = hourValue
        Syne_date_taskHourMap[Date] = TaskDayHour_dict
    return Syne_date_taskHourMap




inputExcel = "C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\Syne Jan Timesheet.xlsx"
inputFormat = "C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\Input_Format.xlsx"
# empId_NameMap = EmpId_Name_Mapping(inputExcel)
# empIdNameProjectMapping = EmpId_Name_Project_Mapping(inputExcel,'321')
# All_dates = get_AllDates_FromSyne(inputExcel)
# EmpId_TaskHour = Resource_Task_TotalHour_mapping(inputExcel,inputFormat,'1234')
# Task_DateHour = Syne_Date_Hours_Mapping(inputExcel,inputFormat, '321', '03-JAN-2021')
#
# print(empId_NameMap)
# print(empIdNameProjectMapping)
# print(All_dates)
# print(EmpId_TaskHour)
# print(Task_DateHour)

