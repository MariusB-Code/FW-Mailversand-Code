import os
import openpyxl # nötig für pandas
import pandas as pd

#Function: checkFilePath: check if the Path to the given file is correct
#--------------------------------------------------------------
def checkFilePath(filePath:str):
    if not os.access(filePath,os.F_OK):
        #print(f"\033[1;37;41m Dateipfad Falsch! Prüfen, ob Dateipfad von {fileType}-Datei richtig ist\033[0;0m")
        return False
    else:
        #print(f"\033[0;32;40m Überprüfung des {fileType}-Dateipfad erfolgreich\033[0;0m")
        return True


#Function: checkFilePaths: check if the file paths for Attachements or embedded images are correct
#--------------------------------------------------------------
def checkArrFilePaths(filePathArr:list):
    i=0
    for filePath in filePathArr:
        if not os.access(filePath,os.F_OK):
            #print(f"\033[1;37;41m Dateipfad Falsch! Prüfen, ob Dateipfad von {fileType} {i+1} richtig ist\033[0;0m")
            return False
        i+=1
    #print(f"\033[0;32;40m Überprüfung der {fileType}-Pfade erfolgreich\033[0;0m")
    return True


#Function: checkFields: check if all nessecary fields are correctly filled in, return True if so
#--------------------------------------------------------------
def checkFields(df):
    for i,row in df.iterrows():
        #check whether Instances "Mail" and "Anrede" contain something
        if (not isinstance(row["Mail"],str)):
            #print(f"\033[1;37;41m Mail-Adresse ist kein Sting! Prüfen, ob 'Mail' in Zeile {i+2} Inhalt hat\033[0;0m")
            return "Email-Adresse ist kein Sting! \nPrüfen, ob 'Mail' in Zeile %d Inhalt hat" %(i+2)
        if (not isinstance(row["Anrede"],str)):
            #print(f"\033[1;37;41m Anrede ist kein Sting! Prüfen, ob 'Anrede' in Zeile {i+2} Inhalt hat\033[0;0m")
            return "Anrede ist kein Sting! \nPrüfen, ob 'Anrede' in Zeile %d Inhalt hat" %(i+2)
        #check whether mail-addresses contains ASCII-Letters only
        if (not (row["Mail"].isascii())):
            #print(f"\033[1;37;41m Mailadresse in Zeile {i+2} enthält nicht-ASCII-Zeichen! Prüfen, ob alles ASCII-Charaktere sind! \033[0;0m")
            #print(f"\033[0;31;40m Email-Adresse nochmals abtippen, danach sollte das Problem behoben sein\033[0;0m")
            return "Email-Adresse in Zeile %d enthält nicht-ASCII-Zeichen! Prüfen, ob alles ASCII-Chars sind! \nBitte Email-Adresse nochmals manuell abtippen" % (i+2)
        #check whether mail-addresses contains space-character
        if (" " in row["Mail"]):
            #print(f"\033[1;37;41m Mailadresse in Zeile {i+2} enthält leerzeichen! Leerzeichen entfernen\033[0;0m")
            return "Email-Adresse in Zeile %d enthält leerzeichen! \nLeerzeichen entfernen"  %(i+2)
        #check whether mail-addresses contains @-character
        if (not "@" in row["Mail"]):
            #print(f"\033[1;37;41m Mailadresse in Zeile {i+2} enthält kein @-Zeichen! Mailadresse prüfen\033[0;0m")
            return "Email-Adresse in Zeile %d enthält kein @-Zeichen! Bitte prüfen" %(i+2)
        #check wheter cc-adresses (if present) conatains ASCII only, no space-char and @-char
        if (isinstance(row["Cc"],str)): #check if cc is sting (if no cc then python defaults to float)
            if (not (row["Cc"].isascii())):
                #print(f"\033[1;37;41m CC-Adresse in Zeile {i+2} enthält nicht-ASCII-Zeichen! Prüfen, ob alles ASCII-Charaktere sind! \033[0;0m")
                #print(f"\033[0;31;40m CC-Adresse nochmals abtippen, danach sollte das Problem behoben sein\033[0;0m")
                return "CC-Adresse in Zeile %d enthält nicht-ASCII-Zeichen! Prüfen, ob alles ASCII-Charaktere sind!" %(i+2)
            if (" " in row["Cc"]):
                #print(f"\033[1;37;41m CC-Adresse in Zeile {i+2} enthält leerzeichen! Leerzeichen entfernen\033[0;0m")
                return "CC-Adresse in Zeile %d enthält leerzeichen! \nLeerzeichen entfernen" %(i+2)
            if (not "@" in row["Cc"]):
                #print(f"\033[1;37;41m CC-Adresse in Zeile {i+2} enthält kein @-Zeichen! CC-Adresse prüfen\033[0;0m")
                return "C-Adresse in Zeile %d enthält kein @-Zeichen! CC-Adresse prüfen" %(i+2)

    #print(f"\033[0;32;40m Überprüfung der 'Anrede', 'Cc' und 'Mail'-Felder erfolgreich\033[0;0m")
    return True

#Function: htmlHasEmbeddedImg: check if html has embedded img, return True if so
#--------------------------------------------------------------
def htmlHasEmbeddedImg(html_path,):
    #open .htm file, check if string [img exists]

    html=[]
    with open(html_path,'r',encoding="utf-8",errors="ignore") as f:
        html = f.read()
        if (html.find("img")==-1):
            return False
        else:
            return True
