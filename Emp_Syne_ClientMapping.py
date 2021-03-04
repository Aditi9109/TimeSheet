#
#
#
import openpyxl
import datetime
import pandas as pd
def Emp_Syne_Client_Mapping(inputFormat):
    Emp_Syne_Client_Mapping={}

    dfFormatSyne = pd.read_excel(inputFormat, "Name_UserId_Mapping")
    empID_List = [x for x in dfFormatSyne['EMP ID'].tolist() if isinstance(x, int)]
    for empID in empID_List:
        userDetailsMapping = []
        userDetailsMapping.append(EmpID_Mapping(inputFormat,empID))
        Emp_Syne_Client_Mapping[empID]=userDetailsMapping
    return Emp_Syne_Client_Mapping


def EmpID_Mapping(inputFormat, empID):
    Map_Syne_UserName = "SYNECHRON USER NAME"
    Map_Client_UserName = "CLIENT USER NAME"
    Emp_Name_Mapping = {}
    EmpID_Mapping=[]
    dfFormatSyne = pd.read_excel(inputFormat, "Name_UserId_Mapping")
    df_new = dfFormatSyne[(dfFormatSyne['EMP ID'] == empID)]
    Syne_UserName_List = [x for x in df_new['SYNECHRON USER NAME'].tolist() if isinstance(x, str)]
    Client_UserName_List = [x for x in df_new['CLIENT USER NAME'].tolist() if isinstance(x, str)]
    for syne_Name, Client_Name in zip(Syne_UserName_List,Client_UserName_List):
        Emp_Name_Mapping[Map_Syne_UserName]=syne_Name
        Emp_Name_Mapping[Map_Client_UserName]=Client_Name
        EmpID_Mapping.append(Emp_Name_Mapping)
    return EmpID_Mapping

inputFormat = "C:\\Users\\PC\Desktop\\Syne_Timesheet\\Input_Format.xlsx"
empDetails=Emp_Syne_Client_Mapping(inputFormat)
print(empDetails)





