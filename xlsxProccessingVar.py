print("Welcome to the XLSX-File Converter!")
print("with this Tool, you can convert any Excel XLSX-File into a fully functional MYSQL-Table.")
print("Keep in mind, that your Table you want to move your Excel-File to has to be named right after your Excel-File.")
print("Right now, the Tool is only Capable of converting one sheet at once.")
print("to make sure everything is working, please insert first your Username, Password, Host and Database you want to move your Excel-Table to.")
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
import pandas as pd                                                                                                             #Import des Pandas-Moduls
import mysql.connector as mc                                                                                                    #Import des MySQL-Moduls
import tkinter as tk                                                                                                            #
from tkinter import filedialog                                                                                                  #
from pathlib import Path                                                                                                        #
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
root = tk.Tk()                                                                                                                  #
root.withdraw()                                                                                                                 #
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Username = input("Log in as: ")                                                                                                 #Angabe des Usernames
PW = input("Password: ")                                                                                                        #Angabe des Passwortes
HostCon = input("Connect to Host: ")                                                                                            #Angabe des Hosts
DB = input("Connect to Database: ")                                                                                             #Angabe der Datenbank
connection=mc.connect(host=HostCon,                                                                                             #Verbindung zur DB mithilfe der Angaben
                      database=DB,                                                                                              #                
                      user=Username,                                                                                            #                      
                      password=PW)                                                                                              #                
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
db_Info = connection.get_server_info                                                                                            #Einholen von Informationen über die Datenbank
print('Informationen des Servers',db_Info)                                                                                      #Ausgabe der Informationen
cursor = connection.cursor()                                                                                                    #Erlaubniserteilung an Python Datenbankeinträge in MySQL zu schreiben
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print("Please locate your xlsx-file: ")                                                                                         #Angabe des xlsx-Dateipfads
xlsxProc = filedialog.askopenfilename()
df = pd.read_excel(xlsxProc)                                                                                                    #Einlesen in den DataFrame
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
PathSchema = input("Schema of your DB: ")                                                                                       #Angabe des DB-Schemas                                                 
Columnlist = list(df.columns)                                                                                                   #Erstellen einer Liste mit den Spaltennamen der Tabelle 
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
StrColumns = ""                                                                                                                 #
for i in Columnlist:                                                                                                            #
    StrColumns = StrColumns + i + ","   
StrColumns = StrColumns[0:len(StrColumns)-1]                                                                                    #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
for i in range(len(df)):
    Values = "("
    for x in Columnlist:                                                                                                        #
        Values = Values +"\""+ str(df[x][i]) +"\","                                                                             #
    cursor.execute('insert into '+PathSchema+"."+Path(xlsxProc).stem+"("+StrColumns+") Value "+Values[0:len(Values)-1]+");")    # 
connection.commit()                                                                                                             #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print('Done')                                                                                                                   #

