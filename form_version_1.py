import tkinter as tk
from tkinter import  filedialog , messagebox    
import sys
import threading
import os
import time
import subprocess



####################################################### REDIRECT OUTPUT CLASS ###############################################################
class RedirectOtput : 

    def __init__(self,text_widget):
        self.text_widget=text_widget

    def write(self,text):
        self.text_widget.insert(tk.END,text)
        self.text_widget.see (tk.END)

    def flush(self):
        pass
#############################################################################################################################################




######################################################### START BROWSER SESSIONS CODE ########################################################
def start_chrome_session(debugging_string):
    chrome_path = "C:\\Program Files\\Google\\Chrome\\Application"
    os.environ["PATH"] = os.pathsep + chrome_path
    time.sleep(0.5)
    
    # Ensure the debugging command is properly defined
    cmd = debugging_string  # Example: "chrome.exe --remote-debugging-port=9222 --user-data-dir='C:\\selenium\\'"
    
    if not cmd:
        print("Debugging mode command is empty. Check the configuration.\n")
        return

    try:
        # Use subprocess.Popen to start Chrome in the background
        process = subprocess.Popen(cmd, shell=True)
        print(f"Chrome session started with PID: {process.pid}\n")
        get_employee_urls()
    except Exception as e:
        print(f"Error starting Chrome session: {e}\n")
    

def start_edge_session(debugging_string):
    edge_path="c:\\program files (x86)\\Microsoft\\Edge\\Application"

    os.environ["PATH"]+=os.pathsep + edge_path
    time.sleep(0.5)
    # Ensure the debugging command is properly defined
    cmd = debugging_string  # Example: "chrome.exe --remote-debugging-port=9222 --user-data-dir='C:\\selenium\\'"
    
    if not cmd:
        print("Debugging mode command is empty. Check the configuration.\n")
        return

    try:
        # Use subprocess.Popen to start Chrome in the background
        process = subprocess.Popen(cmd, shell=True)
        print(f"Edge session started with PID: {process.pid}\n")
        get_employee_urls()
    except Exception as e:
        print(f"Error starting Chrome session: {e}\n")
##########################################################################################################################################


def browsFile(): 
    input_path = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx")])

    if input_path:
        path_box.delete(0,tk.END)
        path_box.insert(0,input_path)


def run_script ():

    browserChoice= browser_var.get()
    input_path =path_box.get()


    if not browserChoice:
        messagebox.showerror("Error","Please Select Browser Type!")
        return
    
    if not input_path : 
        messagebox.showerror("Error","Please Enter Employee Text File Path")
        return
    

    def main_processing_code():
        try :
            print('Starting Program .... \n')
            with open("config.cfg","r",encoding="UTF-8") as cfg:
                debugging_mode_string= cfg.readline().split(",") #the first line in Config file contains the whole debugging string from which chrome and edge debugging mode are extracted
                port=cfg.readline().split(",")[1]
                print (f"{debugging_mode_string[0]}\n{debugging_mode_string[1]}\n{debugging_mode_string[2]}\n {port}")
                cfg.close()
        

            if  browserChoice.lower()=="chrome":
                    start_chrome_session(debugging_string=debugging_mode_string[1])
            elif browserChoice.lower()=="edge":
                    start_edge_session(debugging_string=debugging_mode_string[2])
    
        except Exception as e: 
            messagebox.showerror("Error" , e)
            print(f"{e}\n")

    threading.Thread(target=main_processing_code).start()


