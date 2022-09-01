import email
import os             
import pyautogui
import time
from pandas import read_excel
from time import sleep
from datetime import datetime


try:  
    my_sheet = 'Sheet1' # change it to your sheet name, you can find your sheet name at the bottom left of your excel file
    file_name = 'Email.xlsx' #excel file name should in same place as the index.py
    my_msg='Message you want to Send to everyone in teams'
    
    df = read_excel(file_name, sheet_name = my_sheet)
    email=df["Email ID"]
    print(email)

    # open MS Teams application
    os.startfile("path to your ms teams folder") 
    sleep(2)

    # settings
    for i in email:
        pyautogui.hotkey('ctrl', 'n') # Start a new Chat
        sleep(2)
        pyautogui.typewrite(i)
        sleep(2)
        pyautogui.hotkey('enter') 
        sleep(2)
        pyautogui.hotkey('alt','shift', 'c') # Navigate to Chatiing Option
        sleep(2)
        pyautogui.typewrite(my_msg)
        sleep(2)
        pyautogui.hotkey('enter')
except Exception as e:
    print(e,"EROR")
