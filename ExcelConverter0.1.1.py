print("Welcome to the XLSX-File Converter!")
print("with this Tool, you can convert any Excel XLSX-File into a fully functional MYSQL-Table.")
print("Keep in mind, that your Table you want to move your Excel-File to has to be named right after your Excel-File.")
print("Right now, the Tool is only Capable of converting one sheet at once.")
print("to make sure everything is working, please insert first your Username, Password, Host and Database you want to move your Excel-Table to.")
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
import pandas as pd                                                                                                             # Importieren verschiedener Module: Pandas zum Einlesen der XLSX-Datei
import mysql.connector as mc                                                                                                    # MySQLConnetor um eine Verbindung zu MySQL und der Workbench aufzubauen
from mysql.connector import errorcode                                                                                           # Ausserdem die errorcodes zur Fehleranalyse und Weiterverwendung
import tkinter as tk                                                                                                            # tkinter um einen Popup-Dialog um den Dateipfad der XLSX-Datei aufzurufen
from tkinter import filedialog                                                                                                  # 
from pathlib import Path                                                                                                        #
import sys                                                                                                                      #
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
root = tk.Tk()                                                                                                                  #
root.attributes("-topmost",True)                                                                                                #
root.lift()                                                                                                                     #
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
cursor = connection.cursor()                                                                                                    #
cursor.execute ("SELECT VERSION()")                                                                                             #
row = cursor.fetchone()                                                                                                         #
print("Server Version:", row[0])                                                                                                #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print("Please locate your xlsx-file: ")                                                                                         #
xlsxProc = filedialog.askopenfilename()                                                                                         #
df = pd.read_excel(xlsxProc)                                                                                                    #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
PathSchema = input("Schema of your DB: ")                                                                                       #                                                 
Columnlist = list(df.columns)                                                                                                   #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
try:                                                                                                                            #
    cursor.execute("select"+"*"+"from "+PathSchema+"."+Path(xlsxProc).stem+";")                                                 #
except mc.Error as err:                                                                                                         #
    print("An Error occured: ",err)                                                                                             #
    if err.errno == errorcode.ER_BAD_DB_ERROR:                                                                                  # 
        print("Theres no existing "+PathSchema+"-Schema")                                                                       #
        errBadDB = input("Do you want to create a new Standard-Schema as "+PathSchema+"? (yes/no): ")                           #
        while errBadDB != "yes" and errBadDB != "no":                                                                           #
            errBadDB = input("Input was not detected, please try again (yes/no): ")                                             #
        else:                                                                                                                   #
            pass                                                                                                                #
        if errBadDB == "yes":                                                                                                   #
            print("Creating a new "+PathSchema+"-Schema with standard settings")                                                #
            try:                                                                                                                #
                cursor.execute("create Schema "+PathSchema)                                                                     #
                connection.commit()                                                                                             #
            except mc.Error as err:                                                                                             #
                print("An Error occured: "+err+" try restarting the Converter")                                                 #
        if errBadDB == str("no"):                                                                                               #
            print("Try Adjusting your inputs and restart the Converter")                                                        #
            Stop = input("Press enter to Exit the Converter")                                                                   #
            if Stop == "" and Stop != "":                                                                                       #
                sys.exit()                                                                                                      #
     if err.errno == errorcode.ER_BAD_TABLE_ERROR:                                                                              #
        print("Theres no existing Table named "+Path(xlsxProc).stem)                                                            #
        errBadTable = input("Do you want to create a new Table as "+Path(xlsxProc).stem+" within the "+PathSchema+"?")          #
          while errBadTable != "yes" and errBadTable != "no":                                                                   # 
            errBadTable = input("Input was not detected, please try again (yes/no): ")                                          #
        else:                                                                                                                   #
            pass                                                                                                                #
        if errBadTable = "yes":                                                                                                 #
          print("Creating a new Table \""+Path(xlsxProc).stem+"\" within the "+PathSchema+" Schema")                            #
          try:                                                                                                                  #
            Columnstring = "("+Columnlist[0]+" varchar(99) primary key unique,"                                                 #
            for Title in Columnlist:                                                                                            #
              if Tile = Columnlist[0]:                                                                                          #
                pass                                                                                                            #
              else:                                                                                                             #
                Columnstring += Title+" varchar(99),"                                                                           #
            Columnstring = Columnstring[0:len(Columnstring)-1]                                                                  #
            Columnstring += ");"                                                                                                #
            connection.commit()                                                                                                 #         
            cursor.execute("create table "+PathSchema+"."+Path(xlsxProc).stem+Columnstring                                      #
        if errBadTable == "no":                                                                                                 #
            print("Try Adjusting your inputs and restart the Converter")                                                        #
            Stop = input("Press enter to Exit the Converter")                                                                   #
            if Stop == "" and Stop != "":                                                                                       #
                sys.exit()                                                                                                      #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
StrColumns = ","                                                                                                                #
for i in Columnlist:                                                                                                            #
    StrColumns = StrColumns + i + ","                                                                                           #
StrColumns = StrColumns[1:len(StrColumns)-1]                                                                                    #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
for i in range(len(df)):                                                                                                        #
    Values = "("                                                                                                                #
    for x in Columnlist:                                                                                                        #
        Values = Values +"\""+ str(df[x][i]) +"\","                                                                             #
    cursor.execute('insert into '+PathSchema+"."+Path(xlsxProc).stem+"("+StrColumns+") Value "+Values[0:len(Values)-1]+");")    # 
connection.commit()                                                                                                             #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print('Done')                                                                                                                   #
