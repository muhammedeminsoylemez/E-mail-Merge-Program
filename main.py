from imap_tools import MailBox, AND
from PyPDF2 import PdfMerger
import os
import datetime
import customtkinter
import tkcalendar
import time



switch = True
customtkinter.set_appearance_mode("Dark")

def Appearance():
    global switch
    if switch == False:
        customtkinter.set_appearance_mode("Dark")
        switch = True
    else:
        customtkinter.set_appearance_mode("Light")
        switch = False


customtkinter.set_default_color_theme("green")

root = customtkinter.CTk()
root.geometry("500x350")
root.title("E-mail Merger")


Basarili = 0
Basarisiz = 0
Zerobyte = 0
b = 0
merger = PdfMerger()
lst = []
usercriteria = ""
posta = ""
generatedPassword = ""
userdate = []
kriter = ""
desktop = ""
date = customtkinter.StringVar()
check = customtkinter.StringVar()
seen = True


def Variables():
    global Basarili
    global Basarisiz
    global Zerobyte
    global b
    global merger
    global lst
    global usercriteria
    global posta
    global generatedPassword
    global userdate
    global kriter
    global desktop
    global date
    global check

    Basarili = 0
    Basarisiz = 0
    Zerobyte = 0
    b = 0
    merger = PdfMerger()
    lst = []
    usercriteria = ""
    posta = ""
    generatedPassword = ""
    userdate = []
    kriter = ""
    desktop = ""

def Check():
    global check
    global seen
    if check.get() == "1":
        seen = False

    elif check.get() == "0":
        seen = True


sol_ust_frame = customtkinter.CTkFrame(master=root,
                                          width=500,
                                          height=350,
                                          corner_radius=10)

sol_ust_frame.place(relx=0.0266, rely=0.04, relwidth=0.66666, relheight=0.92)

entry = customtkinter.CTkEntry(master=sol_ust_frame, placeholder_text="Kimden gelen pdfleri indirmek istiyorsunuz")
entry.place(relx=0.050, rely=0.03125, relwidth=0.90, relheight=0.10)

entry2 = customtkinter.CTkEntry(master=sol_ust_frame, placeholder_text="E-Postanızı Giriniz")
entry2.place(relx=0.050, rely=0.15625, relwidth=0.90, relheight=0.10)

entry3 = customtkinter.CTkEntry(master=sol_ust_frame, placeholder_text="Google tarafından oluşturulan şifreyi giriniz")
entry3.place(relx=0.050, rely=0.28125, relwidth=0.90, relheight=0.10)

checkbox1 = customtkinter.CTkCheckBox(master=sol_ust_frame, text="Okunmuş E-Mailleri de birleştir", variable=check, command=Check)
checkbox1.place(relx=0.050, rely=0.40625, relwidth=0.90, relheight=0.10)

label = customtkinter.CTkLabel(master=sol_ust_frame, text="Process:")
label.place(relx=0.125, rely=0.5525, relwidth=0.75, relheight=0.10)

label2 = customtkinter.CTkLabel(master=sol_ust_frame, text="")
label2.place(relx=0.01, rely=0.65125, relwidth=0.98, relheight=0.10)

label3 = customtkinter.CTkLabel(master=sol_ust_frame, text="")
label3.place(relx=0.01, rely=0.77625, relwidth=0.98, relheight=0.10)

label4 = customtkinter.CTkLabel(master=sol_ust_frame, text="")
label4.place(relx=0.01, rely=0.89125, relwidth=0.98, relheight=0.10)


sag_ust_frame = customtkinter.CTkFrame(master=root,
                                          width=500,
                                          height=350,
                                          corner_radius=10)
sag_ust_frame.place(relx=0.6996, rely=0.04, relwidth=0.26666, relheight=0.2)


sag_orta_frame = customtkinter.CTkFrame(master=root,
                                           width=500,
                                           height=350,
                                           corner_radius=10)
sag_orta_frame.place(relx=0.6996, rely=0.26, relwidth=0.26666, relheight=0.41)

date_chooser = tkcalendar.DateEntry(master=sag_orta_frame, width=12, bg="blue", locale="tr_TR", textvariable=date)
date_chooser.place(relx=0.125, rely=0.15, relwidth=0.75, relheight=0.20)

date_label = customtkinter.CTkLabel(master=sag_orta_frame, width=15, text="Tarih Bekleniyor")
date_label.place(relx=0.07, rely=0.65, relwidth=0.85, relheight=0.30)


sag_alt_frame = customtkinter.CTkFrame(master=root,
                                          width=500,
                                          height=350,
                                          corner_radius=10)
sag_alt_frame.place(relx=0.6996, rely=0.68, relwidth=0.26666, relheight=0.28)

merge_label = customtkinter.CTkLabel(master=sag_alt_frame, width=15, text="Merger Bekleniyor")
merge_label.place(relx=0.03, rely=0.6, relwidth=0.94, relheight=0.20)

dark_light_switch = customtkinter.CTkSwitch(sag_alt_frame, text=" Dark/Light", command=lambda: Appearance())
dark_light_switch.place(relx=0.1, rely=0.2, relwidth=0.80, relheight=0.20)


#Functions

def Usercriteria():
    global usercriteria
    usercriteria = entry.get()

def Posta():
    global posta
    posta = entry2.get()

def GeneratedPassword():
    global generatedPassword
    generatedPassword = entry3.get()

def Userdate():
    global userdate
    global date
    z, y, x = date.get().replace(".", " ").split()
    x = int(x)
    y = int(y)
    z = int(z)

    userdate = (datetime.date(x, y, z))

    Change_M_0()
    date_label.configure(text="{}".format(date.get()))

def Kriter():
    global userdate
    global usercriteria
    global kriter
    global seen
    if entry.get() == "":
        kriter = AND(date_gte=userdate, flagged=seen)
    else:
        kriter = AND(from_="'" + usercriteria + "'", date_gte=userdate, flagged=seen)

def Desktop():
    global desktop
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    desktop = desktop.replace(os.sep, '/')

def createPdfsFolder():
    global b
    if 'Pdfs' in os.listdir('{}'.format(desktop)):
        pass
    else:
        os.mkdir('{}/Pdfs/'.format(desktop))
        b = 1

def emailProcess():
    global Basarili
    global Basarisiz
    global Zerobyte
    global lst
    global kriter

    for msg in mailbox.fetch(kriter):
        for att in msg.attachments:
            if att.content_type == 'application/pdf':
                if att.size > 10:
                    label3.configure(text="")
                    root.update()
                    time.sleep(0.05)
                    with open('{}/Pdfs/{}'.format(desktop, att.filename), 'wb') as f:
                        f.write(att.payload)
                        m = att.filename
                        if att.filename in lst:
                            pass
                        else:
                            lst.append(m)
                    label3.configure(text="Dosya Kayıt İşlemi Başarılı")
                    root.update()
                    time.sleep(0.05)
                    Basarili += 1

                else:
                    label3.configure(text="")
                    root.update()
                    time.sleep(0.05)
                    label3.configure(text="Bulunan dosyalar düşük boyutlu olduğu için indirilmedi")
                    root.update()
                    time.sleep(0.05)
                    Zerobyte += 1

            else:
                label3.configure(text="")
                root.update()
                time.sleep(0.05)
                label3.configure(text="Bulunan Dosya Pdf Değil sonraki dosyaya geçiliyor ")
                root.update()
                time.sleep(0.05)
                Basarisiz += 1

def createFinalPdfs():
    global merger

    finalpath = '{}/Final/'.format(desktop)
    a = 1
    if 'Final.pdf' in os.listdir(finalpath):
        while True:
            if 'Final{}.pdf'.format(a) in os.listdir(finalpath):
                a = a + 1
                continue
            else:
                merger.write('{}/Final/Final{}.pdf'.format(desktop, a))
                break

    else:
        merger.write('{}/Final/Final.pdf'.format(desktop))

    merger.close()

def mergePdfsandcreateFinal():
    global merger

    path = '{}/Pdfs/'.format(desktop)

    for file_name in lst:
        merger.append(path + file_name)
    if 'Final' in os.listdir('{}/'.format(desktop)) and Basarili > 0:
        createFinalPdfs()
    elif 'Final' not in os.listdir('{}/'.format(desktop)) and Basarili > 0:
        os.mkdir('{}/Final/'.format(desktop))
        createFinalPdfs()
    else:
        merge_label.configure(text="Uygun Dosya Yok!")

def deletePdfs():

    pdfpath = '{}/Pdfs/'.format(desktop)

    if b == 1:
        for f in os.listdir(pdfpath):
            os.remove(os.path.join(pdfpath, f))
        os.rmdir(pdfpath)
    else:
        for f in lst:
            os.remove(os.path.join(pdfpath, f))

def Change_M_0():
    merge_label.configure(text="Merger Bekleniyor")

def Change_S():
    global Basarili
    label2.configure(text="{} Başarılı".format(Basarili))

def Change_U():
    global Basarisiz
    label3.configure(text="{} Başarısız".format(Basarisiz))

def Change_N():
    global Zerobyte
    label4.configure(text="{} Kullanılamayan".format(Zerobyte))

def Change_M():
    merge_label.configure(text="Merge İşlemi Başarılı")

def Change_L():
    label2.configure(text="")
    label3.configure(text="")
    label4.configure(text="")

def MailBoxWith():
    global posta
    global generatedPassword
    global mailbox
    with MailBox('imap.gmail.com').login(posta, generatedPassword) as mailbox:
        createPdfsFolder()
        emailProcess()

def Main():
    global mailbox

    Change_L()
    root.update()

    Usercriteria()
    Posta()
    GeneratedPassword()
    Userdate()
    Desktop()
    Kriter()
    MailBoxWith()
    mergePdfsandcreateFinal()
    deletePdfs()
    Change_N()
    Change_S()
    Change_U()
    Variables()
    time.sleep(1)
    Change_M()

date_button = customtkinter.CTkButton(sag_orta_frame, text="Set Date", command=lambda: Userdate())
date_button.place(relx=0.125, rely=0.40, relwidth=0.75, relheight=0.20)

start_button = customtkinter.CTkButton(sag_ust_frame, text=" Start", command=lambda : Main())
start_button.place(relx=0.125, rely=0.125, relwidth=0.75, relheight=0.75)
root.mainloop()
