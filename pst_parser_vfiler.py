import win32com.client
import os
import re
from datetime import date

todays_date = date.today()
current_user = os.getlogin()
global divider
divider = "="*65+"\n"
global content
content = f"vFiler rapport skapad: {todays_date} av {current_user}\n\n{divider}"

input("Startat, tryck enter för att fortsätta") # acknowledge start, press enter to continue
path_pst_file = input("Skriv/kopiera in fullständig sökväg till .pst filen: ") # Paste path to .pst file

def find_pst_folder(OutlookObj, pst_filepath) :
    for Store in OutlookObj.Stores :
        if Store.IsDataFileStore and Store.FilePath == pst_filepath :
            return Store.GetRootFolder()
    return None

def enumerate_folders(FolderObj) :
    for ChildFolder in FolderObj.Folders :
        enumerate_folders(ChildFolder)
    iterate_messages(FolderObj)

def iterate_messages(FolderObj) :
    for msg in FolderObj.Items :

        global content
        utfall_rapport = "Error, should not occur"
        total_occurences = 0
        servernamn = "Error, should not occur"
        kund = "Error, should not occur"

        content_from_mail = msg.Body
        get_servernamn = msg.Subject
        servernamn = re.findall('_([\w]+)', get_servernamn)
        find_occurences = re.findall('(\s{3}cifs)|_(cifs01)', content_from_mail)
        total_occurences = len(find_occurences)
        kund = re.findall("\s([a-z]*)_{1}", msg.Subject)
        kund = ' '.join(kund)
        if total_occurences < 25 :
            utfall_rapport = "vFiler rapport inte utan anmmärkning, felkontroll och åtgärd i ärende: XXXXXX "
        else : 
            utfall_rapport = "vFiler rapport utan anmärkning"
        utskrift = f"Kontrollerat vFiler kopior åt {kund.upper()}.\nAntal dagar kopior sparade på: {servernamn} är: ['{total_occurences}']\n{utfall_rapport}"
        content = f"\n{content}\n{utskrift}\n\n{divider}"
        content = content.lstrip()

def print_to_file() :
    create_file = open(os.getcwd() + f"\\vFiler_rapport_{todays_date}.txt", "x")
    create_file.write(content)
    create_file.close

def send_email() :                                                          # Kommer bara köras om kommentar tas bort på rad 73 / remove comment line 83 to use mail
    skicka_mail = win32com.client.Dispatch('Outlook.Application')
    mail = skicka_mail.CreateItem(0)                                        # Avsändare kommer bli Outlook kontot som är igång på datorn. / sender is current user
    mail.To = 'användare@domän'                                             # Hårdkodad mailadress kan läggas till här / add hard coded mail here
    mail.Subject = f"vFiler rapport {todays_date}"
    mail.Body = f"Automatiskt skapat mail, rapport bifogad\n\n{content}"    # Här kan var content skickas om man vill ändra till mail istället för att skapa fil
    attachment = os.getcwd() + f'\\vFiler_rapport_{todays_date}.txt'        # samma rapport som skapats och lagts till i filkatalogen, se funktion print_to_file()
    mail.Attachments.Add(attachment)
    mail.Send()

Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
pst = path_pst_file
Outlook.AddStore(pst)
PSTFolderObj = find_pst_folder(Outlook,pst)
try :
    enumerate_folders(PSTFolderObj)
    print_to_file()
    # send_email()
except Exception as exc :
    print(exc)
finally :
    Outlook.RemoveStore(PSTFolderObj)
