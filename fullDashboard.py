import os
import logging
import time

import datetime as dt
from datetime import timedelta

import win32com.client
from shutil import move

from cryptography.fernet import Fernet
from hash import *

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.support.ui import Select

from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from utils import *

#Set chromedriver options
options = Options()
options.add_experimental_option("detach", True)
options.add_argument('--ignore-certificate-errors')
options.add_argument("--start-maximized")


#Set global variables
start_time = time.time()
yesterday = (dt.datetime.now() - timedelta(days=1)).strftime("%d")
wait = WebDriverWait(chrome, 300)
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
f = Fernet(key)


#Clean up Downloads Folder
folder = "C:\\Users\\Warehouse-MGR\\Downloads"
extension = ".csv"

# traverse the folder
for root, dirs, files in os.walk(folder):
    # loop through the files
    for file in files:
        # check if the file extension matches
        if file.endswith(extension) and root == folder:
            # delete the file
            os.remove(os.path.join(root, file))


#Gather data with selenium
shipheroLogin(str(f.decrypt(email2))[2:-1], str(f.decrypt(pass1))[2:-1])
getPendingShipments()
getProductLocations()
getProducts()
emailShipments()
homebaseLogin(str(f.decrypt(email2))[2:-1], str(f.decrypt(pass2))[2:-1])
getSchedule()
getTimesheets()
getShipments(str(f.decrypt(email1))[2:-1], str(f.decrypt(pass3))[2:-1])
quitChrome()


#Distribute data to shared driver
moveFiles()


#Update workbooks and export PDF
updateData()
updateDashboard()

xl.Quit()


#Send dashboard
slackDashboard(slack_key, "C03U3Q0SKUM")


#Display Calculation time
end_time = time.time()
total_time=end_time-start_time
print("Completed dashboard in {} seconds".format(total_time))
