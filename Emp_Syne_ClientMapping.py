#
#
#
import openpyxl
import datetime
import pandas as pd

from Syne_TestReportMapping import EmpId_Name_Mapping


def Emp_Syne_Client_Mapping(inputFormat,empID):
    Emp_Syne_Client_Mapping={}
    Emp_Syne_Client_Mapping[empID]=EmpID_Mapping(inputFormat,int(empID))
    return Emp_Syne_Client_Mapping


def EmpID_Mapping(inputFormat, empID):
    Map_Syne_UserName = "SYNECHRON USER NAME"
    Map_Client_UserName = "CLIENT USER NAME"
    Emp_Name_Mapping = {}
    dfFormatSyne = pd.read_excel(inputFormat, "Name_UserId_Mapping")
    df_new = dfFormatSyne[(dfFormatSyne['EMP ID'] == empID)]
    Syne_UserName_List = [x for x in df_new['SYNECHRON USER NAME'].tolist() if isinstance(x, str)]
    Client_UserName_List = [x for x in df_new['CLIENT USER NAME'].tolist() if isinstance(x, str)]
    for syne_Name, Client_Name in zip(Syne_UserName_List,Client_UserName_List):
        Emp_Name_Mapping[Map_Syne_UserName]=syne_Name
        Emp_Name_Mapping[Map_Client_UserName]=Client_Name
    return Emp_Name_Mapping

inputFormat = "C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\Input_Format.xlsx"
inputExcel = "C:\\Users\\aditi\\OneDrive\\Desktop\\Vishal_Syne\\Syne Jan Timesheet.xlsx"
# empDetails=Emp_Syne_Client_Mapping(inputFormat,1234)
# print(empDetails)





