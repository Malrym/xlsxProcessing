#---------------------------------------------------------------------------------------------------------------------------
import pandas as pd                                                                                                    
import mysql.connector as mc  
#---------------------------------------------------------------------------------------------------------------------------
Username = input("Log in as: ")
PW = input("Password: ")
HostCon = input("Connect to Host: ")
DB = input("Connect to Database: ")
connection=mc.connect(host=HostCon,
                      database=DB,                                                                               
                      user=Username,                                                                                   
                      password=PW)                                                                               
#---------------------------------------------------------------------------------------------------------------------------
db_Info = connection.get_server_info                                                                                   
print('Informationen des Servers',db_Info)
cursor = connection.cursor() 
#---------------------------------------------------------------------------------------------------------------------------
xlsxProc = input("Path of your xlsx-file: ")
df = pd.read_excel(xlsxProc)
#---------------------------------------------------------------------------------------------------------------------------
PathSchema = input("Schema of your DB: ")
PathTable = input("Table of your DB: ")
LengthColumn = int(input("Number of Columns of your table: "))
Counter = 0
Columnlist = []
while Counter < LengthColumn:
    Column = input("Your "+str(Counter+1)+". Column: ")
    Columnlist.append(Column)
    Counter = Counter + 1
#---------------------------------------------------------------------------------------------------------------------------
StrColumns = ""
Counter = 0
while Counter < LengthColumn:
    StrColumns = StrColumns +Columnlist[Counter]+","
    Counter = Counter+1
StrColumns = StrColumns[0:len(StrColumns)-1]
#---------------------------------------------------------------------------------------------------------------------------
StrValues = ""
Counter = 0
while Counter < LengthColumn:
    StrValues = StrValues + StrValues
    Counter = Counter + 1
#---------------------------------------------------------------------------------------------------------------------------
Counter = 0
CounterExe = 0
Values ="("
while Counter < LengthColumn:
    Values = Values + str(df[Columnlist[Counter]][CounterExe])+"," #Work in Progress
    Counter = Counter + 1
Values = Values[0:len(Values)-1]
Values = Values + ");"
#---------------------------------------------------------------------------------------------------------------------------    
while Counter<len(df):                              
    cursor.execute('insert into '+PathSchema+"."+PathTable+"("+StrColumns+") Value "+Values)
    CounterExe = CounterExe + 1  
connection.commit()    
#---------------------------------------------------------------------------------------------------------------------------
print('done')

