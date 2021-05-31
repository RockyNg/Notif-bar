#!/usr/bin/env python
#gmail
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from pathlib import Path
#Outlook
import win32com.client
import pyttsx3
import outlook
#frame
import tkinter as tk
from tkinter import *
import sys
import io
import tkinter.messagebox
import webbrowser
import time
#import moving file
import shutil

#add logging
import logging


logging.basicConfig(filename='app.log',level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')
def main():
    logging.info("App launched")

    def on_configure(event):
        # update scrollregion after starting 'mainloop'
        # when all widgets are in canvas
        canvas.configure(scrollregion=canvas.bbox('all'))
        #add social media screen
    def social():
        root.withdraw()
        global screen1
        screen1 = Toplevel(root)
        screen1.geometry("400x400")
        screen1.title("Social Medias")
        Button(screen1, text ="Add Gmail",width="10",height="1", command =tutorial).pack()
        Button(screen1, text ="Add Outlook",width="10",height="1", command =create_outlook).pack()
        Button(screen1, text ="Add Discord",width="10",height="1", command =create_discord).pack()
        Button(screen1, text ="Home",width="10",height="1", command =go_back).pack(anchor = "w", side = "bottom")

        #return to main screen from screen 1
    def go_back():
        screen1.destroy()
        root.deiconify()

        #return to main screen from screen 7
    def go_back2():
        screen7.destroy()
        root.deiconify()

        #return to main screen from screen 3
    def go_back3():
        screen2.destroy()
        root.deiconify()

        #return to main screen from screen 4
    def go_back4():
        screen5.destroy()
        root.deiconify()

        #create credentials.json
    def create_gmail():
        screen1.destroy()
        global screen7
        screen7 = Toplevel(root)
        screen7.geometry("400x400")
        screen7.title("Instruction")
        Label(screen7,text="Instruction to Create Gmail API",bg="grey",width="400",height="2", font=("Calibri",13)).pack()
        link3 = Label(screen7, text="Link to start", fg="blue", cursor="hand2")
        link3.pack()
        link3.bind("<Button-1>", lambda e: callback("https://developers.google.com/gmail/api/quickstart/php"))

        Label(screen7, text="Step1: Click the link above.\n").pack(anchor=W)
        Label(screen7, text="Step2: Select account you want from top right corner.\n").pack(anchor=W)
        Label(screen7, text="Step3: Click \"Enable the Gmail API\" button\n").pack(anchor=W)
        Label(screen7, text="Step4: Agree to the terms of service and click \"Next\".\n").pack(anchor=W)
        Label(screen7, text="Step5: Click \"CREATE\".\n").pack(anchor=W)
        Label(screen7, text="Step6: Download the file and Click \"DONE\".\n").pack(anchor=W)
        Label(screen7, text="Click \"Complete button below once you followed all the steps\"\n").pack(anchor=W)
        Button(screen7, text ="Complete",width="10",height="1", command =move_cred).pack()
        Button(screen7, text ="Home",width="10",height="1", command =go_back2).pack(anchor = "w", side = "bottom")


    def move_cred():
        screen7.destroy()
        global screen2
        screen2 = Toplevel(root)
        screen2.geometry("400x400")
        screen2.title("Confirm download")
        Label(screen2, text="Check your Downloads folder for the credentials.json file.").pack(anchor=W)
        Label(screen2, text="If it exists, please enter your Downloads folder file path below").pack(anchor=W)
        Label(screen2, text="Use the guide below if there are difficulties finding your paths").pack(anchor=W)
        global downloads_path
        global app_path

        downloads_path=StringVar()
        app_path=StringVar()
        Label(screen2, text="").pack()
        Label(screen2, text="Your credentials file path").pack()

        downloads_path_entry=Entry(screen2, textvariable = downloads_path)
        downloads_path_entry.pack()
        Button(screen2, text ="Guide",width="10",height="1", command =guide).pack(anchor = "e")

        Button(screen2, text ="Enter",width="10",height="1", command =move).pack()

        Button(screen2, text ="Home",width="10",height="1", command =go_back3).pack(anchor = "w", side = "bottom")

    def guide():
        from PIL import ImageTk, Image

        #This creates the main window of an application
        window = Toplevel(root)
        window.title("Join")
        window.geometry("680x550")
        window.configure(background='grey')

        path = "guide.png"

        #Creates a Tkinter-compatible photo image, which can be used everywhere Tkinter expects an image object.
        img = ImageTk.PhotoImage(Image.open(path))

        #The Label widget is a standard Tkinter widget used to display a text or image on the screen.
        panel = tk.Label(window, image = img)

        #The Pack geometry manager packs widgets in rows or columns.
        panel.pack(side = "bottom", fill = "both", expand = "yes")

        #Start the GUI
        window.mainloop()

        #move credentials from download folder to user folder
    def move():
        screen2.destroy()
        downloads_path_info = downloads_path.get()

        cred_path = Path(downloads_path_info+"\\credentials.json")

        if cred_path.is_file():
             tkinter.messagebox.showinfo('Success',"Your credentials are stored, now continue on the window that popped up")
             original = cred_path
             target = user+'file\credentials.json'
             shutil.move(original,target)
             create_gmail0()

        else:
            tkinter.messagebox.showinfo('Download Path Error',"The credentials file is not in the Downloads file path you've entered. Please enter the correct Downloads file path.")

    def create_outlook():
        tkinter.messagebox.showinfo("Outlook Added", "Emails from your Outlook app are now showing in this App.")
        screen1.destroy()

        file=open(user+'file\outlook.txt',"w")
        file.write("On")
        file.close()
        ofile()
        root.deiconify()
    def create_discord():
         tkinter.messagebox.showerror("Error", "Discord functionality not implemented yet")
    def log_out():
        os.remove("stay.txt")
        os.execl(sys.executable, sys.executable, *sys.argv)
    def callback(url):
        webbrowser.open_new(url)

    #run gmail notifications
    def create_gmail0():
        started = Path(user+'file\credentials.json')

        if started.is_file():

            create_gmail0.num=1
            token_count = Path(user+'file\\token'+str(create_gmail0.num)+'.pickle')
            while token_count.is_file():
               create_gmail0.num=create_gmail0.num+1
               token_count = Path(user+'file\\token'+str(create_gmail0.num)+'.pickle')
            update_clock2()
            if main.roll==1:
                screen4.destroy()
            root.deiconify()

        else:
            create_gmail()
    def tutorial():
        global screen4
        screen4 = Toplevel(root)
        screen4.geometry("400x400")
        screen4.title("Instruction")
        Label(screen4,text="Choose an account to connect",bg="grey",width="400",height="2", font=("Calibri",13)).pack()
        Label(screen4, text="Step1: Select the gmail account you wish to connect. \n").pack(anchor=W)
        Label(screen4, text="Step2: Click \"Advanced\" towards the bottom left.\n").pack(anchor=W)
        Label(screen4, text="Step3: Select Go to Quickstart.\n").pack(anchor=W)
        Label(screen4, text="Step4: Allow Quickstart to view your emails.\n").pack(anchor=W)
        Label(screen4, text="Step5: Click \"Allow\".\n").pack(anchor=W)
        Button(screen4, text ="Enter",width="10",height="1", command =create_gmail0).pack()


    def settings():
        root.withdraw()
        global screen5
        screen5 = Toplevel(root)
        screen5.geometry("400x400")
        screen5.title("Settings")
        screen5.configure(background='white')
        global label5
        label5=Label(screen5, text ="Theme",width="10",height="1").grid(column=0, row=0)
        Button(screen5, text ="Default",width="10",height="1", command =normal).grid(column=1, row=0)
        Button(screen5, text ="Dark",width="10",height="1", command =dark).grid(column=2, row=0)
        label5=Label(screen5, text ="Disable",width="10",height="1").grid(column=0, row=1)
        label5=Label(screen5, text ="Enable",width="10",height="1").grid(column=0, row=2)

        Button(screen5, text ="Gmail",width="10",height="1", command =no_gmail).grid(column=1, row=1)
        Button(screen5, text ="Outlook",width="10",height="1", command =no_outlook).grid(column=2, row=1)
        Button(screen5, text ="Gmail",width="10",height="1", command =yes_gmail).grid(column=1, row=2)
        Button(screen5, text ="Outlook",width="10",height="1", command =yes_outlook).grid(column=2, row=2)

        Button(screen5, text ="Home",width="10",height="1", command =go_back4).grid(column=0, row=3,sticky="s")
        #disable
    def no_gmail():
        label2.configure(text="-------No Gmail Notifications-------------No Gmail Notifications-------")
    def no_outlook():
        label.configure(text="No Outlook Notifications")

        #enable
    def yes_gmail():
        main.roll=0
        create_gmail0()
    def yes_outlook():
        ofile()
    def dark():
        root.configure(background='black')
        canvas.configure(background='black')
        label2.config(fg="white",bg="black")
        label.config(fg="white",bg="black")
        screen5.configure(background='black')
    def normal():
        canvas.configure(background='white')
        root.configure(background='white')
        label2.config(fg="black",bg="white")
        label.config(fg="black",bg="white")
        screen5.configure(background='white')

    def update_clock2():
        if main.roll==0:
            update_clock2.num=main.count
        else:
             screen1.destroy()

             update_clock2.num= create_gmail0.num
        update_clock22()

        #gmail notif
    def update_clock22():
        num=update_clock2.num
        if started.is_file():
            f = open("stay.txt", "r")
            user=f.read()
            f.close()
        else:
            user=test.login_verify.right_user
        my_file = Path(user+'file\credentials.json')
        input_string2=''
        input_string3=''
        input_string4=''


        while num>0:

            my_token = Path(user+'file\\token'+str(num)+'.pickle')
            old_stdout = sys.stdout # Memorize the default stdout stream
            sys.stdout = buffer = io.StringIO()
            #start of gmail
            from contextlib import contextmanager

            @contextmanager
            def suppress_stdout():
                with open(os.devnull, "w") as devnull:
                    old_stdout = sys.stdout
                    sys.stdout = devnull
                    try:
                        yield
                    finally:
                        sys.stdout = old_stdout
            # If modifying these scopes, delete the file token.pickle.
            SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

            creds = None

            #hide the link to gmail link
            with suppress_stdout():
                if os.path.exists(my_token):
                    with open(my_token, 'rb') as token:
                       creds = pickle.load(token)
                ##If there are no (valid) credentials available, let the user log in.
                if not creds or not creds.valid:
                    if creds and creds.expired and creds.refresh_token:
                        creds.refresh(Request())
                    else:
                        flow = InstalledAppFlow.from_client_secrets_file(
                            my_file, SCOPES)
                        creds = flow.run_local_server(port=0)
                    # Save the credentials for the next run

                    with open(my_token, 'wb') as token:
                        pickle.dump(creds, token)

                service = build('gmail', 'v1', credentials=creds)

            results = service.users().messages().list(userId='me', labelIds=['UNREAD', 'INBOX']).execute()
            messages = results.get('messages', [])
            i=0
            message_count=10
            if not messages:
                print('No messages found.')
            else:

                for message in messages[:message_count]:
                    msg= service.users().messages().get(userId='me', id=message['id']).execute()
                    payload = msg['payload']
                    headers = payload['headers']
                    for d in headers:
                        if d['name'] == 'Subject':
                            subject = d['value']
                        if d['name'] == 'From':
                            sender = d['value']
                        if d['name'] == 'To':
                            mail_account = d['value']
                    if i==0:
                        print ("Notifications from \'"+mail_account+"\' account.\n")
                        i=1
                    print("Subject: ", subject, "\n from", sender)

                    print("\n")
                    #time.sleep(2)

            print("-----------------------------------------\n")

            sys.stdout = old_stdout # Put the old stream back in place
            if num==1:
                input_string2 = buffer.getvalue() # Return a str containing the entire contents of the buffer.

                char_list = [input_string2[j] for j in range(len(input_string2)) if ord(input_string2[j]) in range(65536)]
                input_string2=''
                for j in char_list:
                    input_string2=input_string2+j

                logging.info("Gmail alerts: "+ input_string2)
                print(input_string2)
                label2.configure(text=input_string2+input_string3+input_string4)
            if num==2:
                input_string3 = buffer.getvalue() # Return a str containing the entire contents of the buffer.

                char_list = [input_string3[j] for j in range(len(input_string3)) if ord(input_string3[j]) in range(65536)]
                input_string3=''
                for j in char_list:
                    input_string3=input_string3+j

                logging.info("Gmail alerts: "+ input_string3)
                print(input_string3)
                label2.configure(text=input_string4+input_string3)
            if num==3:
                input_string4 = buffer.getvalue() # Return a str containing the entire contents of the buffer.

                char_list = [input_string4[j] for j in range(len(input_string4)) if ord(input_string4[j]) in range(65536)]
                input_string4=''
                for j in char_list:
                    input_string4=input_string4+j

                logging.info("Gmail alerts: "+ input_string4)
                print(input_string4)
                label2.configure(text=input_string4)
            root.after(1100000, create_gmail0)#1000=1 second
            num=num-1
        #try:
    logging.info("Running Outlook code")
    started = Path("stay.txt")

    if started.is_file():

        f = open("stay.txt", "r")
        user=f.read()
        f.close()
    else:
        import test
        user=test.login_verify.right_user

    root = tk.Tk()
    root.geometry("700x450")
    root.configure(background='white')
    #tk.messagebox.showinfo('App','App launched')
    canvas = tk.Canvas(root,width=500, height=500)
    canvas.pack(side=tk.LEFT)

    scrollbar = tk.Scrollbar(root, command=canvas.yview)
    scrollbar.pack(side=tk.LEFT, fill='y')

    canvas.configure(yscrollcommand = scrollbar.set)

    # update scrollregion after starting 'mainloop'
    # when all widgets are in canvas
    canvas.bind('<Configure>', on_configure)

    # --- put frame in canvas ---

    frame = tk.Frame(canvas)
    canvas.create_window((0,0), window=frame, anchor='nw')
    my_file2 = Path(user+'file\outlook.txt')
    label = tk.Label(frame, text="No Outlook Notification", borderwidth=2, relief="solid")
    def ofile():
        label.pack(side="bottom")
        old_stdout = sys.stdout # Memorize the default stdout stream
        sys.stdout = buffer = io.StringIO()
        outlook.main()
        sys.stdout = old_stdout # Put the old stream back in place
        input_string5 = buffer.getvalue() # Return a str containing the entire contents of the buffer.

        char_list = [input_string5[j] for j in range(len(input_string5)) if ord(input_string5[j]) in range(65536)]
        input_string5=''
        for j in char_list:
            input_string5=input_string5+j

        logging.info("outlook alerts: "+ input_string5)
        print(input_string5)
        label.configure(text=input_string5,width="54")
    if my_file2.is_file():
        ofile()
    label2 = tk.Label(frame, text="-------No Gmail Notifications-------------No Gmail Notifications-------", borderwidth=2, relief="solid")
    label2.pack(side="bottom")
    link2 = tk.Label(root,text="Go to Gmail",width="20",height="1", fg="blue", cursor="hand2")
    link4 = tk.Label(root,text="Go to Outlook",width="20",height="1", fg="blue", cursor="hand2")
    text = tk.Label(root,text="hello", fg="blue", cursor="hand2")
    Button(root, text ="Add Social Media",width="20",height="1", command =social).pack()
    Button(root, text ="Settings",width="20",height="1", command =settings).pack()
    Button(root, text ="Log Out",width="20",height="1", command =log_out).pack()
    link2.bind("<Button-1>", lambda e: callback("https://mail.google.com/"))
    link4.bind("<Button-2>", lambda e: callback("https://outlook.office365.com/mail/inbox"))
    link4.pack(side="bottom")
    link2.pack(side="bottom")
    logging.info("Start retrieving Gmail alerts")
    main.roll=1
    my_file = Path(user+'file\credentials.json')
    if my_file.is_file():
        main.count=1
        main.roll=0
        token_count = Path(user+'file\\token'+str(main.count)+'.pickle')
        while token_count.is_file():
           main.count=main.count+1
           token_count = Path(user+'file\\token'+str(main.count)+'.pickle')

        main.count=main.count-1
        create_gmail0()

    main.roll=1
    logging.info("Successfully retrieved Gmail alerts")
    root.mainloop()

main()