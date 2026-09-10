""""Built with Tkinter Designer https://visualtkinter.com/designer"""
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import messagebox
import openpyxl # nötig für pandas
import pandas as pd
import os
import time

import mailversandSMTPhandler
import msgAuxDataCheck
import mailMerger
import msgPacker

#Globale Konstanten
#-------------------------------------------------------------
MAXBURSTMAILS = 100
RATELIMITCODES = [421, 450, 451, 550, 554]

#Globale Variablen
#-------------------------------------------------------------
connTest = ""
saveHostName = ""
savePort = ""
saveUsername = ""
savePassword = ""
on_txtAttachFiles_drop_FIRST=True
on_txtAttackImg_drop_FIRST=True

import tkinter as tk

class SampleText:
    #Sampletext-class für tk.Entry und tk.Text
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        widget.bind("<FocusIn>", self.on_entry_click, add="+")
        widget.bind("<FocusOut>", self.on_focusout, add="+")
        self.on_focusout(None)  # Funktion on_focusout bei Initialisierung aufrufen

    def _get_current_text(self):
        if isinstance(self.widget, tk.Entry):
            return self.widget.get()
        elif isinstance(self.widget, tk.Text):
            return self.widget.get("1.0", "end-1c")
        return ""

    def on_entry_click(self, event):
        if self._get_current_text() == self.text:
            if isinstance(self.widget, tk.Entry):
                self.widget.delete(0, "end")
            elif isinstance(self.widget, tk.Text):
                self.widget.delete("1.0", tk.END)

            self.widget.config(fg='black')

    def on_focusout(self, event):
        if self._get_current_text() == '':
            if isinstance(self.widget, tk.Entry):
                self.widget.insert(0, self.text)
            elif isinstance(self.widget, tk.Text):
                self.widget.insert("1.0", self.text)

            self.widget.config(fg='grey')

class ToolTip:
    """Small hover tooltip. tkinter has no built-in one."""

    def __init__(self, widget, text, delay=250):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip = None
        self._job = None
        # add="+" so this never replaces bindings the widget already has.
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._job = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._job is not None:
            self.widget.after_cancel(self._job)
            self._job = None

    def _show(self):
        if self.tip is not None:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip = tk.Toplevel(self.widget)
        # No title bar or border — it should look like a tooltip, not a window.
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self.tip,
            text=self.text,
            justify="left",
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            padx=6,
            pady=3,
        ).pack()

    def _hide(self, _event=None):
        self._cancel()
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None

class Application:
    def __init__(self, root):
        self.root = root
        root.title("fusion&xpéditeur v1.0")
        root.geometry("520x300")
        root.resizable(False, False)
        

        # Menüleiste
        #-------------------------------------------------------------
        menubar = tk.Menu(root)

        self.datei_menu = tk.Menu(menubar, tearoff=0)
        #self.datei_menu.add_command(label="laden", command=self.on_datei_laden)
        #self.datei_menu.add_command(label="Speichern", command=self.on_datei_speichern)
        self.datei_menu.add_separator()
        self.datei_menu.add_command(label="Beenden", accelerator="Alt+F4", command=self.on_datei_beenden)
        menubar.add_cascade(label="Datei", menu=self.datei_menu)

        #self.extras_menu

        root.configure(menu=menubar)

        # accelerator= only draws the shortcut text; these make the keys work.
        root.bind_all("<Alt-F4>", lambda e: self.on_datei_beenden())

        # Feld links oben
        #-------------------------------------------------------------
        self.frame_left_top = tk.Frame(root, relief="groove", bd=2)
        self.frame_left_top.place(x=0, y=0, width=260, height=180)

        self.lblTitle = tk.Label(self.frame_left_top, text="Server-Einstellungen")
        self.lblTitle.place(x=50, y=0, width=160, height=20)

        self.lblHostName = tk.Label(self.frame_left_top, text="SMTP-Hostname:", anchor="w")
        self.lblHostName.place(x=10, y=20, width=110, height=20)

        self.entryHostName = tk.Entry(self.frame_left_top)
        self.entryHostName.place(x=10, y=40, width=160, height=20)

        self.lblPort = tk.Label(self.frame_left_top, text="Port:", anchor="w")
        self.lblPort.place(x=180, y=20, width=40, height=20)

        vcmd_entryPort = (root.register(self._isNumber), "%P") #prüft bevor es in das Feld schreibt
        self.entryPort = tk.Entry(self.frame_left_top, validate="key", validatecommand=vcmd_entryPort)
        self.entryPort.place(x=180, y=40, width=65, height=20)

        self.lblUserName = tk.Label(self.frame_left_top, text="Nutzername:", anchor="w")
        self.lblUserName.place(x=10, y=60, width=80, height=20)

        self.lbluserName2 = tk.Label(self.frame_left_top, text="(meist E-Mail)", font=("Helvetica", 6, "italic"))
        self.lbluserName2.place(x=90, y=60, width=80, height=20)

        self.entryUserName = tk.Entry(self.frame_left_top)
        self.entryUserName.place(x=10, y=80, width=235, height=20)

        self.lblUserPwd = tk.Label(self.frame_left_top, text="Passwort:", anchor="w")
        self.lblUserPwd.place(x=10, y=100, width=50, height=20)

        self.entryUserPwd = tk.Entry(self.frame_left_top)
        self.entryUserPwd.place(x=10, y=120, width=235, height=20)

        self.btnConnTest = tk.Button(self.frame_left_top, text="Verbindung testen", command=self.on_btnConnTest)
        self.btnConnTest.place(x=10, y=150, width=110, height=20)

        self.lblSettingsCorrect = tk.Label(self.frame_left_top, text="")
        self.lblSettingsCorrect.place(x=125, y=150, width=130, height=20)

        # Feld links unten
        #-------------------------------------------------------------
        self.frame_left_bott = tk.Frame(root, relief="groove", bd=2)
        self.frame_left_bott.place(x=0, y=180, width=260, height=50)

        self.lblEmailsubject = tk.Label(self.frame_left_bott, text="E-Mail-Betreff:", anchor="w")
        self.lblEmailsubject.place(x=10, y=0, width=90, height=20)

        self.entryEmailSubject = tk.Entry(self.frame_left_bott)
        self.entryEmailSubject.place(x=10, y=20, width=235, height=20)

        # Feld rechts oben
        #-------------------------------------------------------------
        self.frame_right_top = tk.Frame(root, relief="groove", bd=2)
        self.frame_right_top.place(x=260, y=0, width=260, height=230)

        self.lblTitle = tk.Label(self.frame_right_top, text="Dateien")
        self.lblTitle.place(x=50, y=0, width=160, height=20)

        self.entryExcel = tk.Entry(self.frame_right_top)
        self.entryExcel.place(x=10, y=20, width=240, height=20)
        ToolTip(self.entryExcel,"Pfad zur Excel-Datei hier einfügen oder Drag-n-Drop")
        SampleText(self.entryExcel,"Pfad/zur/Excel-Datei.xslx oder Drag-n-Drop")
        #Drag-n-Drop
        self.entryExcel.drop_target_register(DND_FILES)
        self.entryExcel.dnd_bind("<<Drop>>", self.on_entryExcel_drop)

        self.entryDocx = tk.Entry(self.frame_right_top)
        self.entryDocx.place(x=10, y=50, width=240, height=20)
        ToolTip(self.entryDocx,"Pfad zur Docx-Datei hier einfügen oder Drag-n-Drop")
        SampleText(self.entryDocx,"Pfad/zur/Docx-Datei.docx oder Drag-n-Drop")
        #Drag-n-Drop
        self.entryDocx.drop_target_register(DND_FILES)
        self.entryDocx.dnd_bind("<<Drop>>", self.on_entryDocx_drop)

        self.entryHtml = tk.Entry(self.frame_right_top)
        self.entryHtml.place(x=10, y=80, width=240, height=20)
        ToolTip(self.entryHtml,"Pfad zur HTML-Datei hier einfügen oder Drag-n-Drop")
        SampleText(self.entryHtml,"Pfad/zur/Html-Datei.htm oder Drag-n-Drop")
        #Drag-n-Drop
        self.entryHtml.drop_target_register(DND_FILES)
        self.entryHtml.dnd_bind("<<Drop>>", self.on_entryHtml_drop)

        self.lblAttach = tk.Label(self.frame_right_top, text="Anhänge")
        self.lblAttach.place(x=50, y=100, width=160, height=20)

        self.txtAttachFiles = tk.Text(self.frame_right_top,wrap="char")
        self.txtAttachFiles.place(x=10, y=120, width=240, height=45)
        ToolTip(self.txtAttachFiles,"Pfade zu den Anhängen hier einfügen oder Drag-n-Drop (trennen mit neuer Zeile)")
        SampleText(self.txtAttachFiles,"Pfade/zu/den/Anhängen.pdf \noder Drag-n-Drop")
        #Drag-n-Drop
        self.txtAttachFiles.drop_target_register(DND_FILES)
        self.txtAttachFiles.dnd_bind("<<Drop>>", self.on_txtAttachFiles_drop)

        self.txtAttachImg = tk.Text(self.frame_right_top,wrap="char")
        self.txtAttachImg.place(x=10, y=175, width=240, height=45)
        ToolTip(self.txtAttachImg,"Pfade zu den einzubettenden Bildern hier einfügen oder Drag-n-Drop (trennen mit neuer Zeile)")
        SampleText(self.txtAttachImg,"Pfade/zu/Embedded-Bildern.jpg \noder Drag-n-Drop")
        #Drag-n-Drop
        self.txtAttachImg.drop_target_register(DND_FILES)
        self.txtAttachImg.dnd_bind("<<Drop>>", self.on_txtAttachImg_drop)

        # Feld Status- und Kontrollleiste unten
        #-------------------------------------------------------------
        self.frame_bott = tk.Frame(root, relief="groove", bd=2)
        self.frame_bott.place(x=0, y=230, width=520, height=70)

        self.txtStatusBar = tk.Text(self.frame_bott,height=6,wrap="word",state="disabled",bg="lightgrey")
        self.txtStatusBar.place(x=10, y=10, width=390, height=50)

        self.btnSendEmails = tk.Button(self.frame_bott, text="E-mails Senden", command=self.on_btnSendEmails,state="disabled")
        self.btnSendEmails.place(x=410, y=10, width=100, height=20)

        self.progBar = ttk.Progressbar(self.frame_bott, orient="horizontal", mode="determinate")
        self.progBar.place(x=410, y=40, width=100, height=20)

    #-------------------------------------------------------------
    # Interne Funtionen
    #-------------------------------------------------------------
    
    # Progress-Bar
    def _ProgBarUpdates(self,progBarObj,val=None, max=None):
        if (max != None): # update max val
            progBarObj.config(maximum=max)
        if (val != None):
            progBarObj.config(value=val)


    # Knopf-Toggle
    #------------------------------------------------------------- 
    def _btnToggleOrSet(self, btnObj,newState=None):
        if (newState=="") or (newState==None):
            stateWas=btnObj.cget("state")#Prüfe zustand des Btn
            if (stateWas=="disabled"):
                btnObj.config(state="normal")
            else:
                btnObj.config(state="disabled")
        else:
            btnObj.config(state=newState)

    # Prüft die gedrückte Taste und erlaubt nur Zahlen oder Backspace
    #-------------------------------------------------------------
    def _isNumber(self, pressedKey):
        return pressedKey == "" or pressedKey.isdigit()

    # ändert den Text und Farbe eines Label
    #-------------------------------------------------------------
    def _changeLabel(self, lblObj, newTxt, txtColour=None, bgColour=None):
        lblObj.config(text=newTxt)
        self.changeTxtBgColour(lblObj,txtColour,bgColour)

    # ändert den Text und Farbe eines Textfeld
    #-------------------------------------------------------------
    def _changeTextField(self, txtObj, newTxt, txtColour=None, bgColour=None):
        stateWas=txtObj.cget("state")#Prüfe zustand des Feldes
        if(stateWas=="disabled"):# wenn das Feld ausgeschaltet war, einschalten
            txtObj.config(state="normal")

        txtObj.delete("1.0","end") # löscht gesamten Eintrag im Textfeld
        txtObj.insert("end",newTxt)
        #txtObj.config()
        self.changeTxtBgColour(txtObj,txtColour,bgColour)
        txtObj.see(tk.END)
        if stateWas=="disabled": # wenn das Feld ausgeschaltet war, wieder ausschalten
            txtObj.config(state="disabled")

    # hängt an das Textfeld an
    #-------------------------------------------------------------
    def _appendTextField(self, txtObj, appTxt, txtColour=None, bgColour=None):
        stateWas=txtObj.cget("state")#Prüfe zustand des Feldes
        if stateWas=="disabled":# wenn das Feld ausgeschaltet war, einschalten
            txtObj.config(state="normal")

        txtObj.insert("end","\n" + appTxt)
        self.changeTxtBgColour(txtObj,txtColour,bgColour)
        txtObj.see("end")
        if stateWas=="disabled": # wenn das Feld ausgeschaltet war, wieder ausschalten
            txtObj.config(state="disabled")

    # ändert den Text und Farbe eines Entry-Felds
    #-------------------------------------------------------------
    def _changeEntry(self, entrObj, newTxt, txtColour=None, bgColour=None):
        stateWas=entrObj.cget("state")
        if stateWas=="disabled":# wenn das Feld ausgeschaltet war, einschalten
            entrObj.config(state="normal")
        entrObj.delete(first=0,last="end")
        entrObj.insert("end",newTxt)
        entrObj.xview("end")
        self.changeTxtBgColour(entrObj,txtColour,bgColour)

        if stateWas=="disabled": # wenn das Feld ausgeschaltet war, wieder ausschalten
            entrObj.config(state="disabled")

    # ändert die Farbe vom Text im Label, Entry oder Textfeld
    #-------------------------------------------------------------
    def changeTxtBgColour(self,obj,txtColour,bgColour):
        if txtColour:
            obj.config(fg=txtColour)
        if bgColour:
            obj.config(bg=bgColour)

    # entfernt die tuple-überbleibsel { und ' und ' und , und } aus dem String
    #-------------------------------------------------------------
    def removeTupleRemainsStr(self,path):
        path=path.lstrip(path[0]).rstrip(path[-1])
        path=path.lstrip(path[0]).rstrip(path[-1])
        path=path.rstrip(path[-1])
        return path
 
    #-------------------------------------------------------------
    # Aufrufen externer Funtionen
    #-------------------------------------------------------------

    #  Verbindungstest durchführen
    #-------------------------------------------------------------
    def on_btnConnTest(self):
        # globale Variablen deklarieren
        global connTest
        global saveHostName
        global savePort
        global saveUsername
        global savePassword

        hostName = self.entryHostName.get()
        port = self.entryPort.get()
        username = self.entryUserName.get()
        password = self.entryUserPwd.get()
        connTest=False
        #print(f"{hostName}")
        if(hostName=="" or port==""):
            #print("Bitte Daten komplett ausfüllen")
            self._changeLabel(self.lblSettingsCorrect,"Daten unvollständig",txtColour="red")
        else:
            self._changeLabel(self.lblSettingsCorrect,"wird getestet...",txtColour="black")
            self._changeTextField(self.txtStatusBar,"Verbindung zum Server wird getestet...",txtColour="black")
            root.update()
            returnMsg=mailversandSMTPhandler.testSMTPconn(hostName,port,username,password)
            if (returnMsg==True):
                self._changeLabel(self.lblSettingsCorrect,"Verbindung i.O.",txtColour="green")
                self._changeTextField(self.txtStatusBar,"")
                self._btnToggleOrSet(self.btnSendEmails,"normal") #aktiviere Btn für SendEmails
                #bekannt gute Einstellungen als globale Variabeln speichern
                connTest = True
                saveHostName = hostName
                savePort = port
                saveUsername = username
                savePassword = password
            else:
                self._changeLabel(self.lblSettingsCorrect,"Test fehlgeschlagen",txtColour="red")
                self._changeTextField(self.txtStatusBar,returnMsg,txtColour="red")

    # Bestätigung zum E-mail senden einholen
    #-------------------------------------------------------------
    def on_btnSendEmails(self):
        # globale Variablen deklarieren
        global connTest
        global saveHostName
        global savePort
        global saveUsername
        global savePassword

        sendWoSubject=False
        sendWoHTML=False
        hostName = self.entryHostName.get()
        port = self.entryPort.get()
        username = self.entryUserName.get()
        password = self.entryUserPwd.get()
        subject = self.entryEmailSubject.get()
        xlsx_path = self.entryExcel.get()
        docx_path = self.entryDocx.get()
        html_path = self.entryDocx.get()

        # falls kein Betreff vorhanden:
        if (subject==""):
            if (messagebox.askyesno("Kein Betreff","Kein Betreff. Senden wirklich beginnen?",icon="question")):
                sendWoSubject=True
            else: #Nein, nicht ohne Betreff fortfahren -> return
                return False
        #falls letzter Verbindungstest erfolgreich UND Server-Einstellungen ohne erneuten Verbindungstest geändert:
        if (connTest==True) and (saveHostName!=hostName or savePort!=port or saveUsername!=username or savePassword!=password):
            if (messagebox.askyesno("Server-Einstellungen geändert","Server-Einstellungen wurden ohne erneuten Verbindungstest geändert. Trotzdem fortfahren?",icon="warning")):
                pass #Ja, ohne Verbindungstest fortfahren
            else: #Nein, nicht ohne Verbindungstest fortfahren -> return
                return False
        #falls Pfadangaben leer sind
        if (xlsx_path==""):
            messagebox.showerror("Keine xlsx-Pfadangaben","Dateipfad zum xlsx-Dokument mit den Empfänderdaten fehlt. Bitte Dateipfad eingeben oder Drag-n-Drop!",icon="error")
            return False
        elif (docx_path=="\n"):
            messagebox.showerror("Keine docx-Pfadangaben","Dateipfad zum docx-Dokument mit dem E-mail-Text fehlt. Bitte Dateipfad eingeben oder Drag-n-Drop!",icon="error")
        if (html_path=="\n"):
            if (messagebox.askyesno("Keine html-Pfadangaben","Dateipfad zum html-Dokument mit dem HTML-E-mail-Text fehlt. E-Mails ohne HTML-Ansicht senden?",icon="warning")):
                sendWoHTML=True
            else:
                return False
        # letzte Bestätigung vor dem Zusammenbauen der Emails
        if (messagebox.askyesno("Senden bestätigen", "Senden wirklich beginnen?")):
            self.emailhandler(sendWoSubject,sendWoHTML)

    # E-mail-Handler. Holt sich alle Daten, Prüft alle Dateipfade, lässt Email zusammenbauen/versenden
    #-------------------------------------------------------------
    def emailhandler(self,woSubject=False,woHtml=False):
        #alle relevanten Sachen holen
        hostName = self.entryHostName.get()
        port = self.entryPort.get()
        username = self.entryUserName.get()
        password = self.entryUserPwd.get()
        subject = self.entryEmailSubject.get()
        xlsx_path = self.entryExcel.get()
        docx_path = self.entryDocx.get()
        html_path = self.entryHtml.get()
        #Anhänge in ein Arr (list) packen, letzter Eintrag entfernen (da immer "")
        arrAttachFiles_paths = self.txtAttachFiles.get(1.0,"end-1c").split("\n")
        #arrAttachFiles_paths.pop()
        print(arrAttachFiles_paths)
        arrAttachImg_paths = self.txtAttachImg.get(1.0,"end-1c").split("\n")
        #arrAttachImg_paths.pop()
        print(arrAttachImg_paths)

        #prüfe die Korrektheit der Pfade
        #xlsx-Pfad
        if (msgAuxDataCheck.checkFilePath(xlsx_path)):
            self._changeTextField(self.txtStatusBar,"Prüfen des xlsx-Pfad erfolgreich",txtColour="black")
            root.update()
        else:
            self._changeTextField(self.txtStatusBar,"angegebener Pfad zur xlsx-Datei falsch! Bitte korrigieren",txtColour="red")
            return False
        root.update()
        if (msgAuxDataCheck.checkFilePath(docx_path)):
            self._changeTextField(self.txtStatusBar,"Prüfen des docx-Pfad erfolgreich",txtColour="black")
            root.update()
        else:
            self._changeTextField(self.txtStatusBar,"angegebener Pfad zur docx-Datei falsch! Bitte korrigieren",txtColour="red")
            return False
        #html-Pfad (falls vorhanden)
        if (woHtml):
            if (msgAuxDataCheck.checkFilePath(html_path)):
                self._changeTextField(self.txtStatusBar,"Prüfen des html-Pfad erfolgreich",txtColour="black")
                root.update()
            else:
                self._changeTextField(self.txtStatusBar,"angegebener Pfad zur html-Datei falsch! Bitte korrigieren",txtColour="red")
                return False
        #Anhänge-Pfade (falls vorhanden)
        if (arrAttachFiles_paths[0]!=""):
            if (msgAuxDataCheck.checkArrFilePaths(arrAttachFiles_paths)):
                self._changeTextField(self.txtStatusBar,"Prüfen Pfade für die Anhänge erfolgreich",txtColour="black")
                root.update()
            else:
                self._changeTextField(self.txtStatusBar,"angegebene Pfade zu Anhänge falsch! Bitte korrigieren",txtColour="red")
                return False
        #Embedded-HTML-Bilder-Pfade (falls HTML vorhanden & falls im HTML darauf verwiesen)
        if(woHtml==False):
            if(msgAuxDataCheck.htmlHasEmbeddedImg(html_path)):
                if (msgAuxDataCheck.checkArrFilePaths(arrAttachImg_paths)):
                    self._changeTextField(self.txtStatusBar,"Prüfen Pfade für die Embedded-Bilder erfolgreich",txtColour="black")
                    root.update()
                else:
                    self._changeTextField(self.txtStatusBar,"angegebene Pfade zu Embedded-Bildern falsch! Bitte korrigieren",txtColour="red")
                    return False
        #Alle Dateipfade wurden erfolgreich überprüft (falls nötig)
        self._changeTextField(self.txtStatusBar,"Alle Dateipfade erfolgreich geprüft",txtColour="black")
        root.update()
        #Excel-Dok prüfen, falls i.O. Verpacken und senden starten
        df = pd.read_excel(xlsx_path)
        #print(df)
        
        checkFieldReturn = msgAuxDataCheck.checkFields(df)

        if (checkFieldReturn==True): # falls checkFields i.O.
            numRows = df.shape[0]
            numMail = 0
            numSavedMail=0

            #check if a save-file exists
            if (os.path.isfile("./MassenMails_saveRow.txt")):
                with open("./MassenMails_saveRow.txt") as f:
                    numSavedMail = int(f.read())
                self._changeTextField(self.txtStatusBar,"save-Datei erkannt. \nStarte ab Mail-Nr. %d" %numSavedMail,txtColour="black")
                root.update()
            #starte loop
            maxRetry=3 # maximale Anzahl von Sendeversuchen
            initDelay= 3600/MAXBURSTMAILS+1
            self._changeTextField(self.txtStatusBar,"starte Versand von %d E-Mails" %numRows,txtColour="black")
            self._ProgBarUpdates(self.progBar,max=numRows)
            root.update()
            for index, row in df.iterrows():
                numMail+=1
                if (numMail > numSavedMail):
                    anrede = row["Anrede"]
                    recipient = row["Mail"]
                    cc = row["Cc"]
                    # merge plain text and html-versions of email
                    text = mailMerger.mergeTXT(docx_path,"[Anrede]",anrede)
                    html = mailMerger.mergeHTML(html_path,"[Anrede]",anrede)
                    #Email verpacken
                    msg = msgPacker.createEmail(username,recipient,cc,subject,text,html,arrAttachFiles_paths,arrAttachImg_paths)
                    #Email versenden
                    for attempt in range(0,maxRetry):
                        returnMsg = mailversandSMTPhandler.sendEmail(msg,username,password,hostName,port,recipient)
                        if (returnMsg==True): #Mailversand erfoglgreich
                            self._changeTextField(self.txtStatusBar,"%d/%d E-Mails erfogreich versendet" %(numMail,numRows),txtColour="black")
                            self._ProgBarUpdates(self.progBar,val=numMail)
                            root.update()
                            break #direkt aus retry-loop ausbrechen
                        else:
                            if (returnMsg.smtp_code in RATELIMITCODES):
                                if (attempt < maxRetry):
                                    delay=initDelay*(2**attempt) # exponentielles Backoff
                                    self._changeTextField(self.txtStatusBar,"Rate-Limit von %d Mails/h erreicht. \nCode %s,%s \nPausiere %d sekunden" %(MAXBURSTMAILS,returnMsg.smtp_code,returnMsg.smtp_error,delay),txtColour="black")
                                    root.update()
                                    time.sleep(delay)
                                else:
                                    self._changeTextField(self.txtStatusBar,"Versand wegen Fehler abgebrochen. \n Bei nächstem Programmstart wird ab Mail Nr. %d gestartet." %numMail,txtColour="red")
                                    root.update()
                                    #print(f"\033[0;31;40m Falls Fehler wegen 'Too many E-mails' ca. 1h warten und erneut versuchen\033[0;0m")
                                    f=open("./MassenMails_saveRow.txt", "w")
                                    f.write(str(numMail))
                                    return False
                                    #print(f"\033[0;33;40m Anzahl versandter Mails gespeichert: startet nächstes mal von Mail Nr. {numMail}\033[0;0m")

            # code hat bis hierhin überlebt: alle Mails erforlgreich versandt
            #print(f"\033[0;32;40m Alle Mails erfolgreich versendet.\033[0;0m")
            self._changeTextField(self.txtStatusBar,"%d/%d Mails erfolgreich versendet" %numMail,numRows,txtColour="green")
            root.update()
            if os.path.isfile("./MassenMails_saveRow.txt"):
                os.remove("./MassenMails_saveRow.txt")
                #print("save-Datei entfernt")
        else:
            self._changeTextField(self.txtStatusBar,checkFieldReturn,txtColour="red")
            return False

        
    # E-mail versenden
    #-------------------------------------------------------------
    def sendEmail(self,msg):
        print("TODO: E-Mails versenden")
    
    #Laden von Eingaben
    def on_datei_laden(self):
        print("TODO: Datei laden")

    #Speichern der Eingaben
    def on_datei_speichern(self):
        print("TODO: Datei speichern")

    #Beenden des Programms
    def on_datei_beenden(self):
        root.destroy()

    #-------------------------------------------------------------
    # Datei-Drag-n-Drop-Events
    #-------------------------------------------------------------
    def on_entryExcel_drop(self, event):
        xlsxPath = str(self.root.tk.splitlist(event.data))
        #print(f"File path: {xlsxPath} dropped on entryExcel")
        xlsxPath=self.removeTupleRemainsStr(xlsxPath) # tuple-überbleibsel {' und ',} entfernen
        self._changeEntry(self.entryExcel,xlsxPath,txtColour="black") #Pfad in Entry-feld eingeben

    def on_entryDocx_drop(self, event):
        docxPath = str(self.root.tk.splitlist(event.data))
        #print(f"File path: {docxPath} dropped on entryDocx")
        docxPath=self.removeTupleRemainsStr(docxPath) # tuple-überbleibsel {' und ',} entfernen
        self._changeEntry(self.entryDocx,docxPath,txtColour="black") #Pfad in Entry-feld eingeben

    def on_entryHtml_drop(self, event):
        htmlPath = str(self.root.tk.splitlist(event.data))
        #print(f"File path: {htmlPath} dropped on entryHtml")
        htmlPath=self.removeTupleRemainsStr(htmlPath) # tuple-überbleibsel {' und ',} entfernen
        self._changeEntry(self.entryHtml,htmlPath,txtColour="black") #Pfad in Entry-feld eingeben

    def on_txtAttachFiles_drop(self, event):
        #globale Variabeln deklarieren
        global on_txtAttachFiles_drop_FIRST

        attachPathTuple = self.root.tk.splitlist(event.data)
        i=0
        for i in range(len(attachPathTuple)):
            if i==0 and on_txtAttachFiles_drop_FIRST==True:
                self._changeTextField(self.txtAttachFiles,str(attachPathTuple[i]),txtColour="black")
                on_txtAttachFiles_drop_FIRST=False
            else:
                self._appendTextField(self.txtAttachFiles,str(attachPathTuple[i]))

            #print(f"File {i} dropped on textAttachFiles: {attachPathTuple[i]}")

    def on_txtAttachImg_drop(self, event):
        #globale Variabeln deklarieren
        global on_txtAttackImg_drop_FIRST

        attachPathTuple = self.root.tk.splitlist(event.data)
        i=0
        for i in range(len(attachPathTuple)):
            if i==0 and on_txtAttackImg_drop_FIRST==True:
                self._changeTextField(self.txtAttachImg,str(attachPathTuple[i]),txtColour="black")
                on_txtAttackImg_drop_FIRST=False
            else:
                self._appendTextField(self.txtAttachImg,str(attachPathTuple[i]))

            #print(f"File {i} dropped on txtAttachImg: {attachPathTuple[i]}")
    

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = Application(root)
    root.mainloop()
