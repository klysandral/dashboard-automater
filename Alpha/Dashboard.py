#Project Notes:
#Add conditional logic for date selection based on day of the week (Monday)
#improve error handling with try/except blocks


import os
import logging
import time

import datetime as dt
from datetime import timedelta

import win32com.client
from shutil import move

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

from googleapiclient.discovery import build
from oauth2client import file, client


#Clean old csv files from downloads folder
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
'''
DANGER WILL ROBINSON DANGER!!!
#Clean old csv files from Metric Data folder
folder = "G:\\Shared drives\\Metric Data"
extension = ".csv"

# traverse the folder
for root, dirs, files in os.walk(folder):
    # loop through the files
    for file in files:
        # check if the file extension matches
        if file.endswith(extension) and root == folder:
            # delete the file
            os.remove(os.path.join(root, file))
'''

#Set chromedriver options
options = Options()
options.add_experimental_option("detach", True)
options.add_argument('--ignore-certificate-errors')
options.add_argument("--start-maximized")


#Set global variables
start_time=time.time()
yesterday = (dt.datetime.now() - timedelta(days=1)).strftime("%d")


# Create a new instance of the Chrome driver
chrome = webdriver.chrome.webdriver.WebDriver(options=options)

print('Starting...')

wait = WebDriverWait(chrome, 30)

#Navigate to ShipHero
chrome.get('https://app.shiphero.com/dashboard/orders/pending_shipments')

wait.until(lambda driver: driver.find_element(By.NAME,"email"))

print("Logging in to ShipHero...")

#Find the username and password fields
username_field = chrome.find_element(By.NAME, "email")
password_field = chrome.find_element(By.NAME, "password")

wait.until(EC.element_to_be_clickable((By.NAME, "email")))

#Enter username and password
username_field.send_keys("nate.s.rhodes@gmail.com")
password_field.send_keys("sl!mel0rd")

#Submit the form
clickable = chrome.find_element(By.NAME, "submit")
ActionChains(chrome).click(clickable).perform()

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn"))
print("Logged in!")

try:
    chrome.execute_script("arguments[0].parentNode.removeChild(arguments[0])", chrome.find_element(By.ID, "pending_shipments_processing"))
except:
    pass
chrome.find_element(By.CLASS_NAME,"btn").click()

time.sleep(1)

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn").is_enabled())

print("Download Started!")

wait.until(lambda driver: any(f.startswith('pending_shipments') and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))

print("Download Complete!")

#Navigate to Product Locations
chrome.get('https://app.shiphero.com/dashboard/product-locations?location_level=all')

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn"))
wait.until(lambda driver: driver.find_element(By.CLASS_NAME, "paginate_button"))
time.sleep(1)
chrome.find_element(By.CLASS_NAME,"btn").click()

time.sleep(2)

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn").is_enabled())
print("Download Started!")

wait.until(lambda driver: any(f.startswith('product_locations') and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))
print("Download Complete!")


#Navigate to Products
chrome.get('https://app.shiphero.com/dashboard/products?kit=0&build_kit=0&dropship=0&virtual=0')

time.sleep(1)

wait.until(EC.element_to_be_clickable((By.LINK_TEXT,"Export All Rows")))

while True:
    try:
        chrome.find_element(By.LINK_TEXT,"Export All Rows").click()
    except ElementClickInterceptedException:
        pass
    else:
        break

wait.until(lambda driver: driver.find_element(By.LINK_TEXT,"Export All Rows").is_enabled())
print("Download Started!")

wait.until(lambda driver: any(f.startswith('product_table') and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))
print("Download Complete!")


#Navigate to Shipments
chrome.get('https://shipping.shiphero.com/shipments-report/')

#Confirm ShipHero Login
print("Confirming Login...")

wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "a.auth0-lock-social-button.auth0-lock-social-big-button"))

time.sleep(1)

chrome.find_element(By.CSS_SELECTOR, "a.auth0-lock-social-button.auth0-lock-social-big-button").click()

wait.until(lambda driver: driver.find_element(By.CLASS_NAME, "sc-18886227-0"))

chrome.find_element(By.CLASS_NAME, "sc-18886227-0").click()

time.sleep(1)

print("Login confirmed.")

#Filter report to yesterday's shipments
created_at = chrome.find_element(By.NAME, "createdAt")
ActionChains(chrome).click(created_at).perform()

between = Select(created_at)
between.select_by_visible_text("between")

dates = chrome.find_element(By.CSS_SELECTOR, "div.react-datepicker__day.react-datepicker__day--0{}".format(yesterday))
ActionChains(chrome).click(dates).perform()
ActionChains(chrome).click(dates).perform()

chrome.find_element(By.CSS_SELECTOR, "button.sc-18886227-0.buupqE.button.button--kind-primary.button--size-normal.button--has-label").click()

#Send report by email
wait.until(lambda chrome: chrome.find_element(By.CLASS_NAME, "sc-4134bb64-0"))

chrome.find_element(By.CLASS_NAME, "sc-4134bb64-0").click()

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"sc-4134bb64-0").is_enabled())

time.sleep(3)

print("Report Emailed!")


#Navigate to Homebase
chrome.get("https://joinhomebase.com/")

wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.button.simple")))

chrome.find_element(By.CSS_SELECTOR, "a.button.simple").click()

wait.until(lambda driver: driver.find_element(By.ID, "account_login"))

account_login = chrome.find_element(By.ID, "account_login")
account_password = chrome.find_element(By.ID, "account_password")
account_login.send_keys("nate.s.rhodes@gmail.com")
account_password.send_keys("aze@zft2qbv3HNV1ruz")
account_password.send_keys(Keys.ENTER)

print("Logging in to Homebase...")

wait.until(lambda driver: driver.find_element(By.ID, "dashboard-location-dashboard"))

print("Logged in!")

chrome.get(
    "https://app.joinhomebase.com/schedule_builder#day/shift/{}/{}/{}".format(
    dt.datetime.now().strftime("%m"),  
    dt.datetime.now().strftime("%d"), 
    dt.datetime.now().strftime("%Y")))

wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "a.btn.btn-primary-action.schedule-header__button.export-roster"))
chrome.find_element(By.CSS_SELECTOR, "a.btn.btn-primary-action.schedule-header__button.export-roster").click()

print("Download Started!")

wait.until(lambda driver: any(f.startswith('daily-roster') and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))

print("Download Complete!")

chrome.get("https://app.joinhomebase.com/timesheets?endDate={}{}{}{}{}{}{}{}{}{}{}".format( 
    dt.datetime.now().strftime("%m"), 
    f"%2F", 
    (dt.datetime.now() - dt.timedelta(days=1)).strftime("%d"), 
    f"%2F", 
    dt.datetime.now().strftime("%Y"), 
    f"&filter=all&groupBy=role&startDate=", 
    dt.datetime.now().strftime("%m"), 
    f"%2F", 
    (dt.datetime.now() - dt.timedelta(days=1)).strftime("%d"), 
    f"%2F", 
    dt.datetime.now().strftime("%Y")))

wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "button.Button.Button--medium.Button--theme-primary.TimesheetsNavigation__download"))
chrome.find_element(By.CSS_SELECTOR, "button.Button.Button--medium.Button--theme-primary.TimesheetsNavigation__download").click()

wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "button.Button.Button--full-width.Button--medium.Button--theme-primary-purple.DownloadTimesheetsModal__next_button"))
chrome.find_element(By.XPATH, "//input[@name='detailed']").click()
chrome.find_element(By.CSS_SELECTOR, "button.Button.Button--full-width.Button--medium.Button--theme-primary-purple.DownloadTimesheetsModal__next_button").click()

print("Download Started!")

wait.until(lambda driver: any(f.startswith('El Famoso') and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))

print("Download Complete!")


#retrieve csv from download link in email
chrome.get("https://mail.google.com/mail/?ui=html")

wait.until(lambda chrome: chrome.find_element(By.NAME, "identifier"))

gmail = chrome.find_element(By.NAME, "identifier")
gmail.send_keys("nathan.rhodes@el-famoso.com")
time.sleep(1)
gmail.send_keys(Keys.ENTER)

wait.until(EC.element_to_be_clickable((By.NAME, "Passwd")))

passwd = chrome.find_element(By.NAME, "Passwd")
passwd.send_keys("DUX@fad7dxk*cub7zva")
passwd.send_keys(Keys.ENTER)

wait.until(lambda chrome: chrome.find_element(By.XPATH, "//td[contains(text(),'noreply@shiphero.com')]"))
wait.until(EC.element_to_be_clickable((By.XPATH, "//td[contains(text(),'noreply@shiphero.com')]/following-sibling::td/a")))

chrome.find_element(By.XPATH, "//td[contains(text(),'noreply@shiphero.com')]/following-sibling::td/a").click()

wait.until(lambda chrome: chrome.find_element(By.LINK_TEXT, "View report"))
chrome.find_element(By.LINK_TEXT, "View report").click()

print("Download Started!")

wait.until(lambda driver: any(f.startswith(dt.datetime.now().strftime("%Y")) and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))

print("Download Complete!")

time.sleep(1)

chrome.close()

#move files to shared drive

src_path = os.path.join(os.path.expanduser("~"), "Downloads")
dst_path = os.path.join("G:\\Shared drives", "Metric Data")

# get the list of files
files = os.listdir(src_path)

# sort by modification time
files.sort(key=lambda x: os.path.getmtime(src_path+"\\"+x))

# get the most recent files
files_to_move = files[-6:]

# move the files
for file in files_to_move:
    src = os.path.join(src_path, file)
    dst = os.path.join(dst_path, file)
    try:
        move(src, dst)
    except:
        PermissionError
        time.sleep(1)
    print("Moved {} to {}".format(file, dst_path))
    time.sleep(1)  # to avoid the same filename conflicts

# distribute the files
src_path = os.path.join("G:\\Shared drives", "Metric Data")
for file in os.listdir(src_path):
    if file.startswith("pending"):
        move(os.path.join(src_path, file), os.path.join(src_path, "Pending Shipments"))
        print("Moved {} to {}".format(file, os.path.join(dst_path, "Pending Shipments")))
        time.sleep(1)  # to avoid the same filename conflicts
    elif file.startswith("product_locations"):
        move(os.path.join(src_path, file), os.path.join(src_path, "Product Locations"))
        print("Moved {} to {}".format(file, os.path.join(dst_path, "Product Locations")))
        time.sleep(1)  # to avoid the same filename conflicts
    elif file.startswith("product_table"):
        move(os.path.join(src_path, file), os.path.join(src_path, "Current Inventory Table"))
        print("Moved {} to {}".format(file, os.path.join(dst_path, "Current Inventory Table")))
        time.sleep(1)  # to avoid the same filename conflicts
    elif file.startswith("daily-roster"):
        move(os.path.join(src_path, file), os.path.join(src_path, "Today's Schedule"))
        print("Moved {} to {}".format(file, os.path.join(dst_path, "Today's Schedule")))
        time.sleep(1)  # to avoid the same filename conflicts
    elif file.startswith("El Famoso"):
        move(os.path.join(src_path, file), os.path.join(src_path, "Timesheet Test"))
        print("Moved {} to {}".format(file, os.path.join(dst_path, "Timesheet Test")))
        time.sleep(1)  # to avoid the same filename conflicts
    elif file.startswith("shipped_items"):
        move(os.path.join(src_path, file), os.path.join(src_path, "Shipped Items"))
        print("Moved {} to {}".format(file, os.path.join(dst_path, "Shipped Items")))
        time.sleep(1)  # to avoid the same filename conflicts
    elif file.startswith("20{}".format(dt.datetime.now().strftime("%y"))):
        move(os.path.join(src_path, file), os.path.join(src_path, "Shipment Tables"))
        print("Moved {} to {}".format(file, os.path.join(dst_path, "Shipment Tables")))
        time.sleep(1)  # to avoid the same filename conflicts
    else:
        time.sleep(1)

# Open excel files from google shared drive
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True

#Open Pending Shipments
wb = xl.Workbooks.Open(os.path.join('G:\\Shared drives\\Reports\\Dashboard Reports\\', 'Pending Shipments.xlsx'))
xl.DisplayAlerts = False

time.sleep(2)

#Refresh all queries
wb.RefreshAll()

time.sleep(1)

print("Refeshing queries...")

xl.CalculateUntilAsyncQueriesDone()

print("Finished!")

#Save and close the workbook
wb.Save()
wb.Close()

#Open Last 7
wb = xl.Workbooks.Open(os.path.join('G:\\Shared drives\\Reports\\Dashboard Reports\\', 'Last 7.xlsx'))
xl.DisplayAlerts = False

time.sleep(2)

#Refresh all queries
wb.RefreshAll()

time.sleep(1)

print("Refeshing queries...")

xl.CalculateUntilAsyncQueriesDone()

print("Finished!")

#Save and close the workbook
wb.Save()
wb.Close()

#Open Dashboard reference
wb = xl.Workbooks.Open(os.path.join('G:\\Shared drives\\Dashboard Drive', 'Dashboard Reference.xlsx'))
xl.DisplayAlerts = False

time.sleep(2)

#refresh all queries
wb.RefreshAll()

time.sleep(1)

print("Refeshing queries...")

xl.CalculateUntilAsyncQueriesDone()

print("Finished!")

#Save the workbook
wb.Save()

time.sleep(2)

#Refresh queries a second time for self-referencing tables
wb.RefreshAll()

time.sleep(1)

print("Refeshing again...")

xl.CalculateUntilAsyncQueriesDone()

print("Finished!")

#Save the workbook
wb.Save()

#Export to PDF
wb.ExportAsFixedFormat(0, os.path.join(
    'G:\\Shared drives\\Dashboard Drive\\Dashboards', 
    'Dashboard {}{}{}'.format(dt.datetime.now().strftime("%y"),
    dt.datetime.now().strftime("%m"),
    dt.datetime.now().strftime("%d"),
    '.pdf')
))
wb.Close()

#Quit the application
xl.Quit()

# define Slack access token
slack_access_token = "xoxb-52563538771-4691190848999-d407QIobcljJZtnLhXKSvTa5"

# define Slack channel name
channel_id = "C03U3Q0SKUM"

# define PDF file path
file_path = os.path.join(
    "G:\\Shared drives\\Dashboard Drive\\Dashboards", 
    "Dashboard {}{}{}{}".format( 
    dt.datetime.now().strftime("%y"), 
    dt.datetime.now().strftime("%m"), 
    dt.datetime.now().strftime("%d"), 
    '.pdf')
) 

client = WebClient(token=slack_access_token)
logger = logging.getLogger(__name__)

try:
    result = client.files_upload_v2(channel=channel_id, initial_comment=f"Attaching {file_path} to this channel", file=file_path)
    logger.info(result)
except SlackApiError as e:
    logger.error("Error uploading file: {}".format(e))

end_time = time.time()
total_time=end_time-start_time
print("Completed dashboard in {} seconds".format(total_time))