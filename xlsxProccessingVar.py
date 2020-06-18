print("Welcome to the XLSX-File Converter!")
print("with this Tool, you can convert any Excel XLSX-File into a fully functional MYSQL-Table.")
print("Keep in mind, that your Table you want to move your Excel-File to has to be named right after your Excel-File.")
print("Right now, the Tool is only Capable of converting one sheet at once.")
print("to make sure everything is working, please insert first your Username, Password, Host and Database you want to move your Excel-Table to.")
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
import pandas as pd                                                                                                             #
import mysql.connector as mc                                                                                                    #
import tkinter as tk                                                                                                            #
from tkinter import filedialog                                                                                                  #
from pathlib import Path                                                                                                        #
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
root = tk.Tk()                                                                                                                  #
root.withdraw()                                                                                                                 #
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Username = input("Log in as: ")                                                                                                 #
PW = input("Password: ")                                                                                                        #
HostCon = input("Connect to Host: ")                                                                                            #
DB = input("Connect to Database: ")                                                                                             #
connection=mc.connect(host=HostCon,                                                                                             #
                      database=DB,                                                                                              #                
                      user=Username,                                                                                            #                      
                      password=PW)                                                                                              #                
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
db_Info = connection.get_server_info                                                                                            #
print('Informationen des Servers',db_Info)                                                                                      #
cursor = connection.cursor()                                                                                                    #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print("Please locate your xlsx-file: ")                                                                                         #
xlsxProc = filedialog.askopenfilename()                                                                                         #
df = pd.read_excel(xlsxProc)                                                                                                    #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
PathSchema = input("Schema of your DB: ")                                                                                       #                                                 
Columnlist = list(df.columns)                                                                                                   #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
StrColumns = ""                                                                                                                 #
for i in Columnlist:                                                                                                            #
    StrColumns = StrColumns + i + ","                                                                                           #
StrColumns = StrColumns[0:len(StrColumns)-1]                                                                                    #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
for i in range(len(df)):                                                                                                        #
    Values = "("                                                                                                                #
    for x in Columnlist:                                                                                                        #
        Values = Values +"\""+ str(df[x][i]) +"\","                                                                             #
    cursor.execute('insert into '+PathSchema+"."+Path(xlsxProc).stem+"("+StrColumns+") Value "+Values[0:len(Values)-1]+");")    # 
connection.commit()                                                                                                             #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print('Done')                                                                                                                   #

