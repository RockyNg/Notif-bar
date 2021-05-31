
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from pathlib import Path

from tkinter import *
import webbrowser
import os
import sys
import time
import io

#Outlook
import win32com.client
import pyttsx3

#import connections
#add logging
import logging

def delete3():
    screen4.destroy()
def delete4():
    screen5.destroy()
def go_back():
    screen1.destroy()
    screen.deiconify()
def go_back_login():
    screen2.destroy()
    screen.deiconify()
def show_me():
    screen18 = Toplevel(screen)
    screen18.geometry("100x100")
    screen18.title("saved")
    Label(screen18, text=login_verify.right_user).pack()
    #Button(screen18, text ="OK",width="10",height="1").pack()
   # messagebox.showinfo('App',login_verify.right_user)

def delete5():
    screen1.destroy()
    screen.deiconify()

def create_notes():
    screen17.destroy()
    global raw_filename
    raw_filename= StringVar()
    global raw_notes
    raw_notes = StringVar()

    screen7 = Toplevel(screen)
    screen7.geometry("400x400")
    screen7.title("info")
    Label(text="Instruction to Create Gmail API",bg="grey",width="400",height="2", font=("Calibri",13)).pack()
    link2 = Label(screen7, text="Link to start", fg="blue", cursor="hand2")

    link2.pack()
    link2.bind("<Button-1>", lambda e: callback("https://developers.google.com/gmail/api/quickstart/python"))
    Label(screen7, text="Step1: Click the link above.\n").pack(anchor=W)
    Label(screen7, text="Step2: Select account you want from top right corner.\n").pack(anchor=W)
    Label(screen7, text="Step3: Click \"Enable the Gmail API\" button\n").pack(anchor=W)
    Label(screen7, text="Step4: Agree to the terms of service and click \"Next\".\n").pack(anchor=W)
    Label(screen7, text="Step5: Click \"CREATE\".\n").pack(anchor=W)
    Label(screen7, text="Step6: Download the file and Click \"DONE\".\n").pack(anchor=W)
    Label(screen7, text="Click \"Complete button below once you followed all the steps\"\n").pack(anchor=W)
    Button(screen7,text="Complete", command = ball).pack()

def social():
    global screen17
    screen17 = Toplevel(screen)
    screen17.geometry("400x400")
    screen17.title("Social Medias")
    Button(screen17, text ="Add Gmail",width="10",height="1", command =create_notes).pack()

def session():
    end()

def password_not_recognised():
    global screen4
    screen4 = Toplevel(screen)
    screen4.geometry("150x100")
    screen4.title("Password failed")
    Label(screen4, text="Password Error").pack()
    Button(screen4, text ="OK",width="10",height="1", command =delete3).pack()
def end():
    screen.destroy()
def delete_notes():

    screen6.destroy()
    screen.deiconify()

def User_not_found():
    global screen5
    screen5 = Toplevel(screen)
    screen5.geometry("150x100")
    screen5.title("Unknown user")
    Label(screen5, text="User Not Found").pack()
    Button(screen5, text ="OK",width="10",height="1", command =delete4).pack()

def register():
    screen.withdraw()
    global screen1
    screen1 = Toplevel(screen)
    screen1.geometry("300x250")
    screen1.title("Register")
    global username
    global password
    global username_entry
    global password_entry
    username=StringVar()
    password=StringVar()

    Label(screen1, text="Please enter details below").pack()
    Label(screen1, text="").pack()
    Label(screen1, text="Username * ").pack()

    username_entry=Entry(screen1, textvariable = username)
    username_entry.pack()

    Label(screen1, text="Password * ").pack()
    password_entry=Entry(screen1, textvariable = password)
    password_entry.pack()

    Label(screen1, text="").pack()
    Button(screen1, text ="Register",width="10",height="1", command =register_user).pack()
    Button(screen1, text ="go back",width="10",height="1", command =go_back).pack()

def register_user():
    username_info = username.get()
    password_info = password.get()

    file=open(username_info,"w")
    file.write(username_info+"\n")
    file.write(password_info)
    file.close()
    if not os.path.exists(username_info+"file"):
        os.mkdir(username_info+"file")
        print("Directory " , username_info+"file" ,  " Created ")
    else:
        print("Directory " , username_info+"file" ,  " already exists")
    username_entry.delete(0, END)
    password_entry.delete(0, END)
    Label(screen1, text="Registration success",fg="green",font=("calibri",11)).pack()
    Button(screen1, text ="OK",width="10",height="1", command =delete5).pack()

def login():
    screen.withdraw()
    global screen2
    screen2 = Toplevel(screen)
    screen2.geometry("300x250")
    screen2.title("Login")

    Label(screen2, text="Please enter details below to login").pack()
    Label(screen2, text="").pack()

    global username_verify
    global password_verify
    global username_entry1
    global password_entry1
    username_verify = StringVar()
    password_verify = StringVar()

    Label(screen2, text="Username * ").pack()
    username_entry1 =Entry(screen2, textvariable = username_verify)
    username_entry1.pack()
    Label(screen2, text="").pack()

    Label(screen2, text="Password * ").pack()
    password_entry1 =Entry(screen2, textvariable = password_verify)
    password_entry1.pack()
    Label(screen2, text="Check the box to stay logged in").pack()
    Checkbutton(screen2, bg="white", command=stay_in).pack()
    Button(screen2, text ="Login",width="10",height="1",command = login_verify).pack()
    Button(screen2, text ="go back",width="10",height="1",command = go_back_login).pack()

def login_verify():
    username1 = username_verify.get()
    password1 = password_verify.get()
    username_entry1.delete(0, END)
    password_entry1.delete(0, END)

    list_of_files=os.listdir()
    if username1 in list_of_files:
        file1 = open(username1,"r")
        verify = file1.read().splitlines()
        if password1 in verify:
            login_verify.right_user= username1
            file = open("stay.txt", "w")
            file.write(username1)
            file.close()
            login_success()
        else:
            password_not_recognised()
    else:
        User_not_found()

def login_success():
    screen2.destroy()
    session()
def stay_in():
    file = open("stay.txt", "w")
    file.write("blank")
    file.close()
def main_screen():
    global screen
    screen =Tk()
    screen.geometry("300x250")
    screen.title("App")
    Label(text="Main Screen",bg="grey",width="300",height="2", font=("Calibri",13)).pack()
    Label(text="").pack()
    Button(text ="Login",width="30",height="2",command = login).pack()
    Label(text="").pack()
    Button(text ="Register",width="30",height="2",command = register).pack()

    screen.mainloop()

main_screen()


