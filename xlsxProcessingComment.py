#Programm zum Importieren von xlsx Dateien in Python
#Auslesen und Importieren nach SQL
#Mithilfe von Anaconda3, Numpy und Pandas
#von Nicolas Csaba Bohocki
#11.06.2020
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
import pandas as pd                                                                                                    #Importieren des Pandas Moduls
import mysql.connector as mc                                                                                           #Importieren des MySQL Moduls  
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- 
Username = input("Username: ")                                                                                         #Angabe des Benutzernames für den Zugriff auf die Datenbank
Password = input("Password: ")                                                                                         #Angabe des Passwortes für den Benutzer zum Zugriff auf die Datenbank
Database = input("Connect to Database: ")
connection=mc.connect(host="DESKTOP-OGGKVHD",
                      database=Database,                                                                               #Verbindungsaufnahme zur Datenbank
                      user=Username,                                                                                   #Verbinden mit der Datenbank
                      password=Password)                                                                               #Einloggen mit dem Userprofil und Passwort
db_Info = connection.get_server_info                                                                                   #Informationen vom Server abrufen
print("Informationen des Servers ",db_Info)                                                                            #Informationen des Servers anzeigen
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
cursor = connection.cursor                                                                                             #Erstellen des Cursors zum Schreiben in MSQL
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
df = pd.read_excel("C:\\Users\\nicol\\OneDrive\\Desktop\\Arbeitsunterlagen\\Nico's\\Anwendungsentwicklung\\Fisch.xlsx") #einlesen der xlsx.datei in den dataframe
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
x = 0
while x<df.len:                                                                                                         #Erstellen einer while-schleife zum eintragen der Werte aus dem Dataframe
    werte = '("'+str(df['Name'][x])+'","'+str(df['Größe'][x])+'",'+str(df['Wert'][x])+')'
    cursor.execute('insert into fisch.fisch(Fischbezeichnung,Größe,Wert) values ' + werte)                              #Ausführen der Schleife zum Schreiben in der MySQL-Tabelle
    x = x+1                                                                                                             #Erhöhen der Variable um 1 


