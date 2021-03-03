
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


def Client_Date_Hours_Mapping(inputExcel,empName,Date):
    Client_Date_Hours_Mapping={}
    TaskDayHour_dict = {}
    client_UserName = []
    Client_date_taskHourMap=[]

    dfClientExcel = pd.read_csv(inputExcel)
    df_new = dfClientExcel[(dfClientExcel['username'] == empName)&(dfClientExcel['local_date'] == Date)]
    client_Hrs= df_new.hours.to_list()
    Client_date_taskHourMap = df_new.jobcode_1.to_list()
    Client_Date_Hours_Mapping[Date] = dict(zip(Client_date_taskHourMap, client_Hrs))
    return Client_Date_Hours_Mapping

def ClientUserNameDetails(inputExcel,empName):
    ClientUserNameDetails = {}
    AllDates=[]
    dfClientExcel = pd.read_csv(inputExcel)
    df_Dates = dfClientExcel[(dfClientExcel['username'] == empName)]
    AllDates= df_Dates.local_date.to_list()
    ClientUserNameDetails[empName] = AllDates
    return ClientUserNameDetails

inputExcel = "C:\\Users\\PC\\Desktop\\Syne_Timesheet\\Client timesheet report daily.csv"
userNames=Client_UserName(inputExcel)
clientSheetDate=get_AllDatesFrom_Client(inputExcel)
ClientUserName_Details=ClientUserNameDetails(inputExcel,'AB.CD@testing.com')
DailyTask_Hour = Client_Date_Hours_Mapping(inputExcel, 'EFG.HIJ@testing.com','01-04-2021')
print(userNames)
print(clientSheetDate)
print(ClientUserName_Details)
print(DailyTask_Hour)


