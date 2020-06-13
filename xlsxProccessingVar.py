#---------------------------------------------------------------------------------------------------------------------------
import pandas as pd                                                                             #Import des Pandas-Moduls                                                                   
import mysql.connector as mc                                                                    #Import des MySQL-Moduls
#---------------------------------------------------------------------------------------------------------------------------
Username = input("Log in as: ")                                                                 #Angabe des Usernames
PW = input("Password: ")                                                                        #Angabe des Passwortes
HostCon = input("Connect to Host: ")                                                            #Angabe des Hosts
DB = input("Connect to Database: ")                                                             #Angabe der Datenbank
connection=mc.connect(host=HostCon,                                                             #Verbindung zur DB mithilfe der Angaben
                      database=DB,                                                              #                
                      user=Username,                                                            #                      
                      password=PW)                                                              #                
#---------------------------------------------------------------------------------------------------------------------------
db_Info = connection.get_server_info                                                            #Einholen von Informationen über die Datenbank
print('Informationen des Servers',db_Info)                                                      #Ausgabe der Informationen
cursor = connection.cursor()                                                                    #Erlaubniserteilung an Python Datenbankeinträge in MySQL zu schreiben
#---------------------------------------------------------------------------------------------------------------------------
xlsxProc = input("Path of your xlsx-file: ")                                                    #Angabe des xlsx-Dateipfads
df = pd.read_excel(xlsxProc)                                                                    #Einlesen in den DataFrame
#---------------------------------------------------------------------------------------------------------------------------
PathSchema = input("Schema of your DB: ")                                                       #Angabe des DB-Schemas
PathTable = input("Table of your DB: ")                                                         #Angabe der DB-Tabelle
LengthColumn = int(input("Number of Columns of your table: "))                                  #Angabe der Anzahl der Spalten der xlsx-Tabelle
Counter = 0                                                                                     #Erstellen der Counter-Variable
Columnlist = []                                                                                 #Erstellen einer leeren Liste
while Counter < LengthColumn:                                                                   #Erstellen einer Schleife um die oben erstellte leere Liste  
    Column = input("Your "+str(Counter+1)+". Column: ")                                         #mit den Spaltennamen 
    Columnlist.append(Column)                                                                   #zu füllen
    Counter = Counter + 1                                                                       #Counter um 1 erhöhen
#---------------------------------------------------------------------------------------------------------------------------
StrColumns = ""                                                                                 #Erstellen einer neuen, leeren Liste
Counter = 0                                                                                     #Counnter zurück auf 0 setzen
while Counter < LengthColumn:                                                                   #Erstellen einer Variable bestehend aus einem für MySQL auslesbarem String 
    StrColumns = StrColumns + Columnlist[Counter] + ","                                         #mit den Spaltennamen der Tabelle
    Counter = Counter + 1                                                                       #Counter um 1 erhöhen
StrColumns = StrColumns[0:len(StrColumns)-1]                                                    #Entfernen des letzten Kommas
#---------------------------------------------------------------------------------------------------------------------------
Counter = 0                                                                                     #Counter zurück auf 0 setzen
CounterCursor = 0                                                                               #Zweiten Counter erstellen und auf 0 setzen
Values ="("                                                                                     #Erstellen einer neuen, leeren Variable
while Counter < LengthColumn:                                                                   #Erstellen einer Schleife
    Values = Values + str(df[Columnlist[Counter]][CounterCursor])+","                           #um diese mit flexiblen Werten zu füllen, die von Pandas auswertbar sind
    Counter = Counter + 1                                                                       #Erhöhen des Zählers um 1 
Values = Values[0:len(Values)-1]                                                                #entfernen des letzten Kommas in der Variable
Values = Values + ");"                                                                          #Abschliessen der Variable, damit diese von MySQL ausgewertet werden kann
#---------------------------------------------------------------------------------------------------------------------------    
CounterCursor = 0                                                                               #Counter zurück auf 0 setzen
while Counter < len(df):                                                                        #Erstellen einer Schleife um
    cursor.execute('insert into '+PathSchema+"."+PathTable+"("+StrColumns+") Value "+Values)    #die Werte in die MySQL-DB eintragen zu können
    CounterCursor = CounterCursor + 1                                                           #Erhöhen des Zählers um 1
connection.commit()                                                                             #Bestätigung zum eintragen der Werte in die DB
#---------------------------------------------------------------------------------------------------------------------------
print('done')

