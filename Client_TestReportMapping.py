
import openpyxl
import datetime
import pandas as pd
def Client_UserName(inputExcel):
    client_UserName = {}
    dfClientExcel = pd.read_csv(inputExcel)
    client_UserName= set(dfClientExcel.username.to_list())
    return client_UserName

def get_AllDatesFrom_Client(inputExcel):
    AllDates = {}
    dfClientExcel = pd.read_csv(inputExcel)
    AllDates= set(dfClientExcel.local_date.to_list())
    return AllDates


def Client_Date_Hours_Mapping(inputExcel,inputFormat,empName,Date):
    Client_Date_Hours_Mapping={}
    TaskDayHour_dict = {}
    client_UserName = []
    Client_date_taskMap=[]
    dfFormatSyne = pd.read_excel(inputFormat, "Client")
    billable = [x for x in dfFormatSyne['Billable'].tolist() if isinstance(x, str)]
    leave = [x for x in dfFormatSyne['Leave'].tolist() if isinstance(x, str)]
    dateColNo = 0
    totalBillableHour= 0
    totalLeaveHour= 0
    dfClientExcel = pd.read_csv(inputExcel)
    df_new = dfClientExcel[(dfClientExcel['username'] == empName)&(dfClientExcel['local_date'] == Date)]
    dailyHourValue= df_new.hours.to_list()
    Client_date_taskMap = df_new.jobcode_1.to_list()
    for taskName , hourValue in zip(Client_date_taskMap,dailyHourValue):
        if taskName in billable:
            Task = "Billable"
            if isinstance(hourValue, int) or isinstance(hourValue, float):
                totalBillableHour = totalBillableHour + hourValue
                TaskDayHour_dict[Task] = totalBillableHour
            elif isinstance(hourValue, str):
                if 'Leave' in TaskDayHour_dict.keys():
                    if isinstance(TaskDayHour_dict[Task], int) or isinstance(TaskDayHour_dict[Task], float):
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
                    if isinstance(TaskDayHour_dict[Task], int) or isinstance(TaskDayHour_dict[Task], float):
                        pass
                else:
                    TaskDayHour_dict[Task] = hourValue
        Client_Date_Hours_Mapping[Date] = TaskDayHour_dict
    return Client_Date_Hours_Mapping

def ClientName_DateHours_Mapping(inputExcel,inputFormat,empName):
    ClientUserNameDetails = {}
    AllDates={}
    Day_Hrs_Mapping=[]
    dfClientExcel = pd.read_csv(inputExcel)
    df_Dates = dfClientExcel[(dfClientExcel['username'] == empName)]
    AllDates= set(df_Dates.local_date.to_list())
    for Hrs_Map in AllDates:
        Day_Hrs_Mapping.append(Client_Date_Hours_Mapping(inputExcel,inputFormat,empName,Hrs_Map))
        ClientUserNameDetails[empName]=Day_Hrs_Mapping
    return ClientUserNameDetails

inputExcel = "C:\\Users\\PC\\Desktop\\Syne_Timesheet\\Client timesheet report daily.csv"
inputFormat = "C:\\Users\\PC\Desktop\\Syne_Timesheet\\Input_Format.xlsx"
#userNames=Client_UserName(inputExcel)
#clientSheetDate=get_AllDatesFrom_Client(inputExcel)
ClientUserName_Details1=ClientName_DateHours_Mapping(inputExcel,inputFormat,'AB.CD@testing.com')
ClientUserName_Details2=ClientName_DateHours_Mapping(inputExcel,inputFormat,'PQR.TIM@testing.com')
ClientUserName_Details3=ClientName_DateHours_Mapping(inputExcel,inputFormat,'EFG.HIJ@testing.com')
#DailyTask_Hour = Client_Date_Hours_Mapping(inputExcel, 'EFG.HIJ@testing.com','01-04-2021')
#print(userNames)
#print(clientSheetDate)
print(ClientUserName_Details1)
print(ClientUserName_Details2)
print(ClientUserName_Details3)
#print(DailyTask_Hour)


