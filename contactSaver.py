#!python3.3
#-------------------------------------------------------------------------------
# Name:        FODAY SAHID CONTACT SAVER SOFTWARE
# Purpose:
#
# Author:      FODAY SAHID
#
# Created:     12/11/2015
# Copyright:   (c) FODAY SAHID 2015
# Licence:     <FODAY SAHID LICENSE>
#-------------------------------------------------------------------------------

from tkinter import *

from tkinter import ttk, filedialog

from tkinter import messagebox

from PIL import ImageTk, Image

import smtplib, win32com.client, sqlite3

speak = win32com.client.Dispatch('Sapi.SpVoice')

volume = 100

rate = 0

speak.Volume = volume

speak.Rate = rate

class FodaySoftware:
    def __init__(self, master):
        self.master = master
        self.master.title("CONTACT SAVER")
        self.master.geometry("400x420")
        self.master.resizable(0,0)
        self.master.iconbitmap("MY_PIC.ico")

        self.addIcon = ImageTk.PhotoImage(Image.open('icons/add.jpg'))
        self.ExitIcon = ImageTk.PhotoImage(Image.open('icons/exit.jpg'))
        self.SaveIcon = ImageTk.PhotoImage(Image.open('icons/save.jpg'))
        self.addIcon = ImageTk.PhotoImage(Image.open('icons/add.jpg'))
        self.viewIcon = ImageTk.PhotoImage(Image.open('icons/view.jpg'))
        self.homeIcon = ImageTk.PhotoImage(Image.open('icons/home.jpg'))
        self.searchIcon = ImageTk.PhotoImage(Image.open('icons/search.jpg'))
        self.emailIcon = ImageTk.PhotoImage(Image.open('icons/email.jpg'))
        self.backupIcon = ImageTk.PhotoImage(Image.open('icons/backup.jpg'))
        self.developerIcon = ImageTk.PhotoImage(Image.open('icons/developer.jpg'))
        self.bgPic = ImageTk.PhotoImage(Image.open('pic.jpg'))
        self.sendIcon = ImageTk.PhotoImage(Image.open('icons/send.jpg'))
        self.developerImg = ImageTk.PhotoImage(Image.open('DeveloperPic.jpg'))

        self.lbl = Label(self.master, image = self.bgPic,width = 400,height = 420)
        self.lbl.place(x = 0, y = 30)

        self.MenuBar = Menu(self.master)

        self.fileMenu = Menu(self.MenuBar, tearoff = 0, bg = "white")
        self.fileMenu.add_command(label = "Home", image = self.homeIcon, compound = LEFT, command = lambda: self.homeCmd(self), accelerator = "Ctrl + H")
        self.fileMenu.add_separator()
        self.fileMenu.add_command(label = "Exit", command = self.master.destroy, image = self.ExitIcon, compound = LEFT, accelerator = "Alt + F4")
        self.MenuBar.add_cascade(label = "File", menu = self.fileMenu)

        self.contactMenu = Menu(self.MenuBar, tearoff = 0, bg = "white")
        self.contactMenu.add_command(label = "Add Contact", command = lambda: self.AddContacts(self), image = self.addIcon, compound = LEFT, accelerator = "Ctrl + A")
        self.contactMenu.add_separator()
        self.contactMenu.add_command(label = "View Contacts", command = lambda: self.viewAll(self), image = self.viewIcon, compound = LEFT, accelerator = "Ctrl + V")
        self.contactMenu.add_separator()
        self.contactMenu.add_command(label = "Search", command = lambda: self.searchRecord(self), image = self.searchIcon, compound = LEFT, accelerator = "Ctrl + S")
        self.contactMenu.add_separator()
        self.contactMenu.add_command(label = "Backup Contacts", command = lambda: self.backupContacts(self), image = self.backupIcon, compound = LEFT, accelerator = "Ctrl + B")
        self.MenuBar.add_cascade(label = "Contacts", menu = self.contactMenu)

        self.developerMenu = Menu(self.MenuBar, tearoff = 0, bg = "white")

        self.developerMenu.add_separator()
        
        self.developerMenu.add_command(label = "Foday S.N Kamara", command = None, image = self.developerIcon, compound = LEFT)
       
        self.MenuBar.add_cascade(label = "Developer", menu = self.developerMenu)

        self.aboutMenu = Menu(self.MenuBar, tearoff = 0, bg = "white")
        self.aboutMenu.add_command(label = "Developer", image = self.developerIcon, compound = LEFT, accelerator = "Ctrl + F", command = lambda:self.AboutMe(self))
        self.aboutMenu.add_separator()
        self.aboutMenu.add_command(label = "Contact Developer", command = lambda: self.emailMe(self), image = self.emailIcon, compound = LEFT, accelerator = "Ctrl + M")
        self.MenuBar.add_cascade(label = "About", menu = self.aboutMenu)


        self.master.config(menu = self.MenuBar)

        self.ShortcutBar = Frame(self.master, height = 30, bg = "white")
        self.ShortcutBar.pack(expand = 0, fill = X)


        self.home_lbl = Label(self.ShortcutBar, image = self.homeIcon,text = "Home", compound = "top", bg = "white", cursor = "hand2")
        self.home_lbl.pack(side = LEFT)
        self.home_lbl.bind('<Button-1>', self.homeCmd)


        self.add_lbl = Label(self.ShortcutBar, image = self.addIcon, text = "Add Contact", compound = "top", bg = "white", cursor = "hand2")
        self.add_lbl.pack(side = LEFT)
        self.add_lbl.bind('<Button-1>', self.AddContacts)

        self.view_lbl = Label(self.ShortcutBar, image = self.viewIcon, text = "View Contacts", compound = "top", bg = "white", cursor = "hand2")
        self.view_lbl.pack(side = LEFT)
        self.view_lbl.bind('<Button-1>', self.viewAll)

        self.search_lbl = Label(self.ShortcutBar, text = "Search", image = self.searchIcon, compound = 'top', bg = "white", cursor = "hand2")
        self.search_lbl.pack(side = LEFT)
        self.search_lbl.bind('<Button-1>', self.searchRecord)


        self.email_lbl = Label(self.ShortcutBar, text = "Contact Developer", image = self.emailIcon, compound = "top", bg = "white", cursor = "hand2")
        self.email_lbl.pack(side = LEFT)
        self.email_lbl.bind('<Button-1>', self.emailMe)

        #STARTING OF ADD CONTACT WIDGET - FODAY S.N KAMARA

        self.lblFrame = LabelFrame(self.master, text = "ADD NEW CONTACT", font = "Nina 12 bold")

        self.nameEntry = StringVar()
        self.phoneEntry = StringVar()

        self.fullNameLbl = Label(self.lblFrame, text = "FULL NAME", font = "Nina 12 bold")

        self.fullNameTxt = Entry(self.lblFrame, textvariable = self.nameEntry)

        self.phoneLbl = Label(self.lblFrame, text = "PHONE NUMBER", font = "Nina 12 bold")

        self.phoneTxt = Entry(self.lblFrame, textvariable = self.phoneEntry)

        self.addContactBtn = Button(self.lblFrame, text = "ADD CONTACT", font = "Arial 12 bold", command = self.InsertRecord)

        self.showBtn = Button(self.master, text = "DISPLAY CONTACTS",  font = "Arial 10 bold", command = self.DisplayRecords)

        self.tree = ttk.Treeview(self.master, columns = ("FULL NAME", "PHONE NUMBER"))
        self.tree.heading("FULL NAME", text = "FULL NAME")
        self.tree.heading("PHONE NUMBER", text = "PHONE NUMBER")
        self.tree.column('FULL NAME', stretch = YES, minwidth = 0, width = 150)
        self.tree.column('PHONE NUMBER', stretch = YES, minwidth = 0, width = 150)
        self.tree.column('#0', stretch = NO, minwidth = 0, width = 0)

        self.scrollbar = Scrollbar(self.tree)


        #CLOSING OF ADD CONTACT WIDGET -- FODAY S.N KAMARA

        #STARTING VIEW CONTACT WIDGET --- FODAY S.N KAMARA

        self.viewTree = ttk.Treeview(self.master, columns = ("FULL NAME", "PHONE NUMBER"))
        self.viewTree.heading("FULL NAME", text = "FULL NAME")
        self.viewTree.heading("PHONE NUMBER", text = "PHONE NUMBER")
        self.viewTree.column('FULL NAME', stretch = YES, minwidth = 0, width = 150)
        self.viewTree.column('PHONE NUMBER', stretch = YES, minwidth = 0, width = 150)
        self.viewTree.column('#0', stretch = NO, minwidth = 0, width = 0)

        self.viewTreeScroll = Scrollbar(self.viewTree)
        self.BackupBtn = Button(self.master, text = "BACKUP CONTACTS", command = lambda: self.backupContacts(self))

        #CLOSING VIEW CONTACT WIDGET --- FODAY S.N KAMARA

        #STARTING OF SEARCH WIDGET

        self.searchEntry = StringVar()
        self.lblSearchFrame = LabelFrame(self.master, text = "SEARCH", font = "Nina 12 bold")
        self.searchNameLbl = Label(self.lblSearchFrame, text = "ENTER NAME", font = "Nina 12 bold")
        self.searchNameTxt = Entry(self.lblSearchFrame, textvariable = self.searchEntry)
        self.searchButton = Button(self.lblSearchFrame, text = "SEARCH", font = "Arial 12 bold", command = self.searchContacts)

        self.searchTree = ttk.Treeview(self.master, columns = ("FULL NAME", "PHONE NUMBER"))
        self.searchTree.heading("FULL NAME", text = "FULL NAME")
        self.searchTree.heading("PHONE NUMBER", text = "PHONE NUMBER")
        self.searchTree.column('FULL NAME', stretch = YES, minwidth = 0, width = 150)
        self.searchTree.column('PHONE NUMBER', stretch = YES, minwidth = 0, width = 150)
        self.searchTree.column('#0', stretch = NO, minwidth = 0, width = 0)

        self.searchScroll = Scrollbar(self.searchTree)


        #CLOSING OF SEARCH WIDGET

        self.master.bind('<Control-KeyPress-M>', self.emailMe)
        self.master.bind('<Control-KeyPress-m>', self.emailMe)
        self.master.bind('<Control-KeyPress-F>', self.AboutMe)
        self.master.bind('<Control-KeyPress-f>', self.AboutMe)
        self.master.bind('<Control-KeyPress-H>', self.homeCmd)
        self.master.bind('<Control-KeyPress-h>', self.homeCmd)
        self.master.bind('<Control-KeyPress-A>', self.AddContacts)
        self.master.bind('<Control-KeyPress-a>', self.AddContacts)
        self.master.bind('<Control-KeyPress-V>', self.viewAll)
        self.master.bind('<Control-KeyPress-v>', self.viewAll)
        self.master.bind('<Control-KeyPress-S>', self.searchRecord)
        self.master.bind('<Control-KeyPress-s>', self.searchRecord)
        self.master.bind('<Control-KeyPress-B>', self.backupContacts)
        self.master.bind('<Control-KeyPress-b>', self.backupContacts)
        self.master.protocol('WM_DELETE_WINDOW', self.sayGoodBye)

    def sayGoodBye(self):
        speak.Speak('THANKS FOR USING THE CONTACT SAVER SOFTWARE GOOD BYE')
        self.master.destroy()

    def emailMe(self, e):
        self.top = Toplevel(self.master)
        self.top.geometry("300x300")
        self.top.title("CONTACT FORM")
        self.top.resizable(0,0)
        self.top.config(bg = "pink")
        self.top.iconbitmap("MY_PIC.ico")
        self.top.transient(self.master)

        self.entry = StringVar()
        self.entry2 = StringVar()
        self.entry3 = StringVar()

        Label(self.top, text = "FIRST NAME", bg = "pink", font = "Nina 12 bold").place(x = 100, y = 0)
        self.Fname = Entry(self.top, textvariable = self.entry)
        self.Fname.place(x = 50, y = 20, width = 200)

        Label(self.top, text = "LAST NAME", bg = "pink", font = "Nina 12 bold").place(x = 100, y = 40)
        self.Lname = Entry(self.top, textvariable = self.entry2)
        self.Lname.place(x = 50, y = 60, width = 200)

        Label(self.top, text = "MESSAGE", bg = "pink", font = "Nina 12 bold").place(x = 100, y = 80)
        self.txtMessage = Text(self.top)
        self.txtMessage.place(x = 50, y = 100, width = 200, height = 150)

        self.sendBtn = Button(self.top, text = 'SEND MAIL', command = self.sendMail, image = self.sendIcon, cursor = 'hand2')
        self.sendBtn.place(x = 80, y = 260)


    def AddContacts(self, e):
        self.lbl.place_forget()
        self.viewTree.place_forget()
        self.BackupBtn.place_forget()
        self.lblSearchFrame.place_forget()
        self.searchTree.place_forget()
        self.lblFrame.place(x = 20, y = 60, width = 350, height = 140)
        self.fullNameLbl.place(x = 0, y = 0)
        self.fullNameTxt.place(x = 140, y = 0, width = 200)
        self.phoneLbl.place(x = 0, y = 40)
        self.phoneTxt.place(x = 140, y = 40, width = 200)
        self.addContactBtn.place(x = 100, y = 80)
        self.showBtn.place(x = 120, y = 210)
        self.tree.config(yscrollcommand = self.scrollbar.set)
        self.scrollbar.config(command = self.tree.yview)
        self.tree.place(x = 10, y = 240, width = 380, height = 150)
        self.scrollbar.pack(side = RIGHT, fill = Y)

    def viewAll(self, e):
            self.lblFrame.place_forget()
            self.tree.place_forget()
            self.showBtn.place_forget()
            self.lblSearchFrame.place_forget()
            self.searchTree.place_forget()
            self.lbl.place_forget()
            self.viewTree.config(yscrollcommand = self.viewTreeScroll.set)
            self.viewTreeScroll.config(command = self.viewTree.yview)
            self.viewTreeScroll.pack(side = RIGHT, fill = Y)
            self.viewTree.place(x = 10, y = 60, width = 380, height = 310)
            self.BackupBtn.place(x = 10, y = 370)

            self.rec = self.viewTree.get_children()
            for data in self.rec:
                self.viewTree.delete(data)
            try:
                self.conn = sqlite3.connect("FODAY SAHID.db")
                self.c = self.conn.cursor()
            except Exception as e:
                speak.Speak("ERROR")
                messagebox.showerror(title = "ERROR", message = str(e))
            else:
                self.data = self.c.execute("SELECT ContactName, ContactNumber FROM CONTACTS ORDER BY ContactName ASC")
                for row in self.data:
                    self.viewTree.insert("", END, values = (row[0], row[1]))
                self.conn.close()


    def searchRecord(self, e):
            self.lblFrame.place_forget()
            self.tree.place_forget()
            self.showBtn.place_forget()
            self.viewTree.place_forget()
            self.BackupBtn.place_forget()
            self.lbl.place_forget()
            self.lblSearchFrame.place(x = 20, y = 60, width = 350, height = 120)
            self.searchNameLbl.place(x = 0, y = 0)
            self.searchNameTxt.place(x = 140, y = 0, width = 200)
            self.searchButton.place(x = 110, y = 50)
            self.searchTree.config(yscrollcommand = self.searchScroll.set)
            self.searchScroll.config(command = self.searchTree.yview)
            self.searchTree.place(x = 10, y = 240, width = 380, height = 150)
            self.searchScroll.pack(side = RIGHT, fill = Y)

    def homeCmd(self, e):
            self.lblFrame.place_forget()
            self.tree.place_forget()
            self.showBtn.place_forget()
            self.viewTree.place_forget()
            self.BackupBtn.place_forget()
            self.lblSearchFrame.place_forget()
            self.searchTree.place_forget()
            self.lbl.place(x = 0, y = 30)

    def sendMail(self):
        self.first_name = str(self.Fname.get())
        self.last_name = str(self.Lname.get())
        self.mailMessage = str(self.txtMessage.get(index1 = 0.0, index2 = 400.100))

        if self.first_name == "" or self.last_name == "" or self.mailMessage == "":
            speak.Speak("ERROR")
            messagebox.showerror(title = "ERROR", message = "MISSING FORM VALUE(S)")
        else:
            try:
                mail = smtplib.SMTP('smtp.mail.yahoo.com', 587)
            except Exception as e:
                speak.Speak("FAILED TO SEND EMAIL")
                messagebox.showerror(title = "ERROR", message = str(e))
            else:
                self.info = self.first_name + " " + self.last_name + '\n' + self.mailMessage
                mail.ehlo()
                mail.starttls()
                mail.login('fodays.nkamara65@yahoo.com', '088767795')
                mail.sendmail('fodays.nkamara65@yahoo.com','fodaysnkamara@gmail.com',self.info)
                mail.close()
                speak.Speak("MAIL SENT SUCCESSFULLY")
                messagebox.showinfo(title = "EMAIL SENT", message = "MESSAGE SENT")

    def AboutMe(self, e):
        self.top = Toplevel(self.master)
        self.top.geometry("150x420")
        self.top.title("ABOUT ME")
        self.top.resizable(0,0)
        self.top.config(bg = "black")
        self.top.iconbitmap("MY_PIC.ico")
        self.top.transient(self.master)

        self.Piclbl = Label(self.top, text = "FODAY\nS.N \nKAMARA\nIS A\nCOMPUTER PROGRAMMER\nWHO SPEND MOST\nOF HIS TIME\nWRITING CODES\nAND DEBUGGING\nPROGRAMS\nemail:\nfodaysnkamara@gmail.com\nCellPhone:\n+232-88-76-77-95", image = self.developerImg, bg = 'black', fg = 'white', compound = TOP)
        self.Piclbl.pack()
    def InsertRecord(self):
        self.FullName = str(self.fullNameTxt.get())
        self.PhoneNumber = str(self.phoneTxt.get())
        if self.FullName == "" or self.PhoneNumber == "":
            speak.Speak("ERROR")
            messagebox.showerror(title = "ERROR", message = "MISSING FORM VALUE(S)")
        else:
            try:
                self.conn = sqlite3.connect("FODAY SAHID.db")
                self.c = self.conn.cursor()
            except Exception as e:
                speak.Speak("ERROR")
                messagebox.showerror(title = "ERROR", message = str(e))
            else:
                self.c.execute("INSERT INTO CONTACTS(ContactName, ContactNumber) VALUES(?,?)", (self.FullName, self.PhoneNumber))
                self.conn.commit()
                self.conn.close()
                speak.Speak("SAVING SUCCESSFUL")
                messagebox.showinfo(title = "SAVE CONTACT", message = "SAVING SUCCESSFUL")
                self.nameEntry.set("")
                self.phoneEntry.set("")

    def DisplayRecords(self):
        self.rec = self.tree.get_children()
        for data in self.rec:
            self.tree.delete(data)
        try:
            self.conn = sqlite3.connect("FODAY SAHID.db")
            self.c = self.conn.cursor()
        except Exception as e:
            speak.Speak("ERROR")
            messagebox.showerror(title = "ERROR", message = str(e))
        else:
            self.data = self.c.execute("SELECT ContactName, ContactNumber FROM CONTACTS ORDER BY ContactName ASC")
            for row in self.data:
                self.tree.insert("", END, values = (row[0], row[1]))
            self.conn.close()

    def backupContacts(self, e):
        try:
            self.conn = sqlite3.connect("FODAY SAHID.db")
            self.c = self.conn.cursor()
        except Exception as e:
            speak.Speak("ERROR")
            messagebox.showerror(title = "ERROR", message = str(e))
        else:
            try:
                self.filename = filedialog.asksaveasfilename(defaultextension = ".txt", filetypes = [("Text Document", "*.txt")])
                self.myFile = open(self.filename, 'w')
            except Exception as e:
                speak.Speak("ERROR")
                messagebox.showerror(title = "ERROR", message = str(e))
            else:
                self.data = self.c.execute("SELECT * FROM CONTACTS")
                for rows in self.data:
                    self.records = str(rows[0]) + " " + str(rows[1]) + " " + str(rows[2]) + '\n'
                    self.myFile.write(self.records)
                self.myFile.close()
                self.conn.close()

    def searchContacts(self):
        try:
            self.conn = sqlite3.connect("FODAY SAHID.db")
            self.c = self.conn.cursor()
        except Exception as e:
            speak.Speak("ERROR")
            messagebox.showerror(title = "ERROR", message = str(e))
        else:
            self.searchName = str(self.searchNameTxt.get())
            if self.searchName == "":
                speak.Speak("ERROR")
                messagebox.showerror(title = "ERROR", message = "MISSING FORM VALUE")
            else:
                self.treeData = self.searchTree.get_children()
                for rows in self.treeData:
                    self.searchTree.delete(rows)
                self.data = self.c.execute("SELECT * FROM CONTACTS WHERE ContactName = '%s'" % self.searchName)
                for rows in self.data:
                    self.searchTree.insert("", END, values = (rows[1], rows[2]))
                self.conn.close()
                self.searchEntry.set("")





def main():
    top = Tk()
    structure = FodaySoftware(top)
    top.mainloop()

if __name__ == "__main__":
    main()

