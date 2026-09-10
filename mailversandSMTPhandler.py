import smtplib
import socket

def testSMTPconn(hostName,port,username,password,use_tls=True):
    try:
        if port == 465: # Port 465 nutzt normalerweise SSL/TLS
            smtp = smtplib.SMTP_SSL(hostName, port, timeout=10)
        else:
            # Port 587 (oder 25) nutzt Standard-SMTP
            smtp = smtplib.SMTP(hostName, port, timeout=10)
            smtp.ehlo()
            if use_tls:
                smtp.starttls()  # Verschlüsselung starten
                smtp.ehlo()
        #print("Server und Port sind erreichbar")

        #Login versuchen
        smtp.login(username, password)
        #print("Benutzername und Passwort sind korrekt")
        
        smtp.quit()
        return True

    except socket.gaierror:
        errorMsg="Fehler: Servername konnte nicht aufgelöst werden (Falscher Hostname?)"
    except (socket.timeout, TimeoutError):
        errorMsg="Fehler: Zeitüberschreitung (Falscher Port oder Firewall blockiert?)"
    except smtplib.SMTPAuthenticationError:
        errorMsg="Fehler: Zugangsdaten falsch (Nutzername oder Passwort falsch?)"
    except smtplib.SMTPException as e:
        errorMsg="SMTP-Fehler: %s" %e
    except Exception as e:
        errorMsg="Allgemeiner Fehler: %s" %e

    #print(errorMsg)
    return errorMsg

# Function sendEmail: sends the email, returns True if successful
#--------------------------------------------------------------
def sendEmail(msg,sender,password,hostName,portNr,recipient):
    try:
        smtp = smtplib.SMTP(hostName, portNr)
        smtp.starttls()
        smtp.login(sender,password)
        smtp.send_message(msg,sender,recipient)
        smtp.quit
        #print(f"message {numMail} of {numRows} sent to: {recipient}, cc:{cc}")
        return True

    except (smtplib.SMTPDataError, smtplib.SMTPResponseException) as e:
        return e

"""
    except smtplib.SMTPDataError as e:
        #print(f"Fehler: Rate-Limit oder Daten. Code: {e.smtp_code}, {e.smtp_error}")
        errorMsg = "Fehler: Rate-Limit oder Daten. \nCode %s, %s" %(e.smtp_code, e.smtp_error)
    except smtplib.SMTPResponseException as e:
        #print(f"SMTP-Fehler vom Server. Code: {e.smtp_code}, {e.smtp_error}")
        errorMsg = "SMTP-Fehler vom Server. \nCode %s, %s" %(e.smtp_code, e.smtp_error)
    except smtplib.SMTPException as e:
        #print(f"SMTP-Fehler: {e}")
        errorMsg = "SMTP-Fehler %s" %e
    except Exception as e:
        #print(f"Allgemeiner Fehler: {e}")
        errorMsg = "Allgemeiner Fehler %s" %e
        
    return errorMsg
"""
