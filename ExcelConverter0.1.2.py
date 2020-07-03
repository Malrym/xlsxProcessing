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
from pathlib import Path                                                                                                        # Importieren des pathlib-Moduls zwecks kompateren Pfadangaben zur Verwendung innerhalb von MySQL-Querys
import sys                                                                                                                      # Importieren des Sys-Moduls um in Bestimmten fällen das Programm zum Beenden zu zwingen 
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
root = tk.Tk()                                                                                                                  # Hier wird der Popup-Dialog zur Pfadauswahl des XLSX-Pfads Definiert
root.attributes("-topmost",True)                                                                                                # Rootattribut "-topmost" soll dafür sorgen, dass das Fenster immer über anderen geöffnet wird
root.lift()                                                                                                                     # Der Dialog wird in den Vordergrund gehoben 
root.withdraw()                                                                                                                 # root wird zurückgezogen
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Username = input("Log in as: ")                                                                                                 # Inputeingaben, um abzufragen, als welcher Nutzer,
PW = input("Password: ")                                                                                                        # mit welchem Passwort,
HostCon = input("Connect to Host: ")                                                                                            # zu welchem Host, 
DB = input("Connect to Database: ")                                                                                             # sich auf welche Datenbank 
connection=mc.connect(host=HostCon,                                                                                             # verbunden werden soll
                      database=DB,                                                                                              # Verbindung wird versucht aufzubauen           
                      user=Username,                                                                                            #                      
                      password=PW)                                                                                              #                
cursor = connection.cursor()                                                                                                    # cursor wird als Variable definiert, um auf MySQL zugreifen zu können
cursor.execute ("SELECT VERSION()")                                                                                             # Es wird in MySQL die Version abgefragt, um sicherzustellen, das eine Verbindung besteht
row = cursor.fetchone()                                                                                                         # Es wird die das Ergebnis der ersten Reihe der Abfrage abgerufen 
print("Server Version:", row[0])                                                                                                # Hier wird das Ergebnis der Abfrage dargestellt
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print("Please locate your xlsx-file: ")                                                                                         # Aufforderung, den Dateipfad der xlsx anzugeben
xlsxProc = filedialog.askopenfilename()                                                                                         # der Dateidialog wird geöffnet und das Ergebnis auf der Variable xlsxProc gespeichert 
df = pd.read_excel(xlsxProc)                                                                                                    # Einlesen der xlsx in den DataFrame(df)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
PathSchema = input("Schema of your DB: ")                                                                                       # Das Schema (bzw. die Datenbank) wird abgefragt und auf der Variable PathSchema gespeichert                                                
Columnlist = list(df.columns)                                                                                                   # Es wird eine Liste der einzelnen Zeilennamen auf der Variable Columnlist gespeichert
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
try:                                                                                                                            # Dieser Bereich soll 2 bestimmte Bugs umgehen und behandeln
    cursor.execute("select"+"*"+"from "+PathSchema+"."+Path(xlsxProc).stem+";")                                                 # Hier wird als erstes eine Abfrage erstellt um zu kontrollieren, ob es zu Fehlermeldungen kommt
except mc.Error as err:                                                                                                         # wenn ein Fehler auftritt wird dieser Teil des Programmes ausgeführt
    print("An Error occured: ",err)                                                                                             # Zuerst wird die ursprüngliche Fehlermeldung ausgegeben
    if err.errno == errorcode.ER_BAD_DB_ERROR:                                                                                  # Falls der Fehler ER_BAD_DB_ERROR auftritt wird der Teil des Programmes ausgeführt
        print("Theres no existing "+PathSchema+"-Schema")                                                                       # Ausgabe, was genau zum Ausführen des Programms fehlt, sprich, das Schema bzw. die Datenbank
        errBadDB = input("Do you want to create a new Standard-Schema as "+PathSchema+"? (yes/no): ")                           # auf der Variable errBadDB wird gespeichert, ob ein neues Schema erstellt werden soll
        while errBadDB != "yes" and errBadDB != "no":                                                                           # Eine Schleife, BIS die Eingabe die gewünschte ist:
            errBadDB = input("Input was not detected, please try again (yes/no): ")                                             # wird zur erneuten Eingabe aufgefordert
        else:                                                                                                                   # Falls die Eingabe einer der beiden gewünschten ist:
            pass                                                                                                                # wird diese Schleife übersprungen
        if errBadDB == "yes":                                                                                                   # Wenn die Eingabe "yes" ist:
            print("Creating a new "+PathSchema+"-Schema with standard settings")                                                # Wird ausgegeben, dass versucht wird ein neues Schema / eine neue Datenbank mit Standardeinstellung, d.h. ohne bestimmte Argumente zu erstellen
            try:                                                                                                                # Es wird versucht:
                cursor.execute("create Schema "+PathSchema)                                                                     # Das neue Schema zu erstellen 
                connection.commit()                                                                                             # und den Befehl auszuführen
            except mc.Error as err:                                                                                             # WENN Fehler auftreten:
                print("An Error occured: "+err+" try restarting the Converter")                                                 # wird Ausgegeben, dass es einen Fehler gab
                Stop = input("Press enter to Exit the Converter")                                                               # und um einen Neustart des Programms gebeten, weitere Bugs wurden noch nicht eingebunden an dieser Stelle
                  if Stop == "" and Stop != "":                                                                                 # Egal welche eingabe getätigt wird, aber auf jeden Fall einmal Enter:
                    sys.exit()                                                                                                  # wird das Programm beendet
        if errBadDB == str("no"):                                                                                               # Falls die Eingabe, die oben auf der Variable errBadDB gespeichert wurde "no" ist:
            print("Try Adjusting your inputs and restart the Converter")                                                        # Wird um die Korrekturen der Eingaben gebeten und
            Stop = input("Press enter to Exit the Converter")                                                                   # der Benutzer gebeten, Enter zu drücken, um das Programm zu beenden und neu zu starten
            if Stop == "" and Stop != "":                                                                                       # auch hier kommt es nicht auf die Eingabe an, aber auf jeden Fall einmal Enter:
                sys.exit()                                                                                                      # um wieder das Programm zum beenden zu zwingen. 
     if err.errno == errorcode.ER_BAD_TABLE_ERROR:                                                                              # Hier wird der nächste Fehler, nämlich ER_BAD_TABLE_ERROR behandelt
        print("Theres no existing Table named "+Path(xlsxProc).stem)                                                            # Es wird der Fehler erneut ausgegeben.
        errBadTable = input("Do you want to create a new Table as "+Path(xlsxProc).stem+" within the "+PathSchema+"? (yes/no) ")# Und wieder Abgefragt. Diesmal allerdings, ob eine neue Tabelle mit den oben angegebenen Werten erstellt werden soll
          while errBadTable != "yes" and errBadTable != "no":                                                                   # und wieder eine Schleife erstellt:
            errBadTable = input("Input was not detected, please try again (yes/no): ")                                          # um erneut zur Eingabe aufgefordert, bis die Eingabe eine gewünschte ist
        else:                                                                                                                   # wenn kein Fehler besteht
            pass                                                                                                                # mache weiter.
        if errBadTable = "yes":                                                                                                 # Wenn die eingabe "yes" gemacht wurde:
          print("Creating a new Table \""+Path(xlsxProc).stem+"\" within the "+PathSchema+" Schema")                            # Zeige, dass probiert wird eine Neue Tabelle im Angegebenen Schema zu erstellen
          try:                                                                                                                  # Versuche:
            Columnstring = "("+Columnlist[0]+" varchar(99) primary key unique,"                                                 # Erstelle einen String mit einem Primärschlüssel, erster Wert der Spalten-Liste
            for Title in Columnlist:                                                                                            # Erstelle eine Schleife für jedes Attribut in der Spalten-Liste
              if Title = Columnlist[0]:                                                                                         # Beim ersten eintrag der Liste:
                pass                                                                                                            # überspringe.
              else:                                                                                                             # Ansonsten
                Columnstring += Title+" varchar(99),"                                                                           # Ergänze den ZeilenString um einen neuen eintrag mit den Attributen "varchar(99)"
            Columnstring = Columnstring[0:len(Columnstring)-1]                                                                  # Entferne das letzte Komma
            Columnstring += ");"                                                                                                # und schliesse mit ");" den String ab, damit MySQL ihn weiterverwenden kann.
            cursor.execute("create table "+PathSchema+"."+Path(xlsxProc).stem+Columnstring                                      # Schreiben des dazugehörigen MySQL-Befehls mithilfe der Variablen.
            connection.commit()                                                                                                 # und bestätigen der Abfrage
        if errBadTable == "no":                                                                                                 # Fall die Antwort "no" ist.
            print("Try Adjusting your inputs and restart the Converter")                                                        # Gebe den Text aus und
            Stop = input("Press enter to Exit the Converter")                                                                   # Erstelle eine Variable zum Beenden des Programms
            if Stop == "" and Stop != "":                                                                                       # Egal was die eingabe ist, mindestens einmal enter
                sys.exit()                                                                                                      # Beende das Programm
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
StrColumns = ","                                                                                                                # Erstelle einen neuen String mit Spaltennamen
for i in Columnlist:                                                                                                            # für jedes Element in der Spalten-Liste
    StrColumns = StrColumns + i + ","                                                                                           # ergänze den String um das Element sowie um ein Komma
StrColumns = StrColumns[1:len(StrColumns)-1]                                                                                    # Streiche das Letzte Komma, nachdem die Schleife durchlaufen wurde.
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
for i in range(len(df)):                                                                                                        # Für jedes Element im Dataframe
    Values = "("                                                                                                                # Erstelle zuerst eine Variable mit dem Namen "Values"
    for x in Columnlist:                                                                                                        # für jedes Element in der Spalten-Liste
        Values = Values +"\""+ str(df[x][i]) +"\","                                                                             # Ergänze den Values-String um den Eintrag an der jeweiligen Position im Dataframe
    cursor.execute('insert into '+PathSchema+"."+Path(xlsxProc).stem+"("+StrColumns+") Value "+Values[0:len(Values)-1]+");")    # Schreibe eine Abfrage in MySQL mit den entsprechenden Angaben bestehend aus Schema/Datenbank, der Tabelle, den Spalten der Tabelle, sowie den aktuell gespeicherten Werten im Values-String
connection.commit()                                                                                                             # Bestätige das Ausführen der Abfrage und somit das übernehmen der Werte der Excel-Tabelle in die MySQL-Datenbank.
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print('Done')                                                                                                                   # Bestätigen des Abschliessen das Programms durch eine Textausgabe. Done.
