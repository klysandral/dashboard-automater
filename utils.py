import os
import logging
import time
import calendar

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
from selenium.common.exceptions import ElementClickInterceptedException, NoSuchElementException
from selenium.webdriver.support.ui import Select

from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

#Set chromedriver options
options = Options()
options.add_experimental_option("detach", True)
options.add_argument('--ignore-certificate-errors')
options.add_argument("--start-maximized")


# Create a new instance of the Chrome driver
chrome = webdriver.chrome.webdriver.WebDriver(options=options)


#Set global variables
today = dt.datetime.today().strftime("%A")
yesterday = (dt.datetime.now() - timedelta(days=1)).strftime("%d")
friday = (dt.datetime.now() - timedelta(days=3)).strftime("%d")
yesterday_name = (dt.datetime.now() - timedelta(days=1)).strftime("%A")
friday_name = (dt.datetime.now() - timedelta(days=3)).strftime("%A")
yesterday_month_name = calendar.month_name[int((dt.datetime.now() - timedelta(days=1)).strftime("%m"))]
friday_month_name = calendar.month_name[int((dt.datetime.now() - timedelta(days=3)).strftime("%m"))]
yesterday_year = (dt.datetime.now() - timedelta(days=1)).strftime("%Y")
friday_year = (dt.datetime.now() - timedelta(days=3)).strftime("%Y")

if 4 <= int(yesterday) <= 20 or 24 <= int(yesterday) <= 30:
    yesterday_suffix = "th"
else:
    yesterday_suffix = ["st", "nd", "rd"][int(yesterday) % 10 - 1]

if 4 <= int(friday) <= 20 or 24 <= int(friday) <= 30:
    friday_suffix = "th"
else:
    friday_suffix = ["st", "nd", "rd"][int(friday) % 10 - 1]

wait = WebDriverWait(chrome, 120)
xl = win32com.client.Dispatch("Excel.Application")


#ShipHero Login function
def shipheroLogin(email, password):
    chrome.get('https://app.shiphero.com/dashboard')

    wait.until(lambda driver: driver.find_element(By.NAME,"email"))

    print("Logging in to ShipHero...")

    #Find the username and password fields
    username_field = chrome.find_element(By.NAME, "email")
    password_field = chrome.find_element(By.NAME, "password")

    wait.until(EC.element_to_be_clickable((By.NAME, "email")))

    #Enter username and password
    username_field.send_keys(email)
    password_field.send_keys(password)

    #Submit the form
    clickable = chrome.find_element(By.NAME, "submit")
    ActionChains(chrome).click(clickable).perform()

    wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn"))
    print("Logged in!")


#ShiphHero login confirm function
def confirmLogin():
    #Confirm ShipHero Login
    print("Confirming Login...")

    wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "a.auth0-lock-social-button.auth0-lock-social-big-button"))

    time.sleep(1)

    chrome.find_element(By.CSS_SELECTOR, "a.auth0-lock-social-button.auth0-lock-social-big-button").click()

    wait.until(lambda driver: driver.find_element(By.CLASS_NAME, "sc-18886227-0"))

    chrome.find_element(By.CLASS_NAME, "sc-18886227-0").click()

    time.sleep(1)

    print("Login confirmed.")

#Function to retrieve Pending Shipments csv file
def getPendingShipments():
    chrome.get('https://app.shiphero.com/dashboard/orders/pending_shipments')
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


#Function to retrieve Product Locations csv file
def getProductLocations():
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


#Function to retrieve Products csv file
def getProducts():
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


#Function to trigger Shipments csv file to be emailed
def emailShipments():
    chrome.get('https://shipping.shiphero.com/shipments-report/')

    confirmLogin()

    created_at = chrome.find_element(By.NAME, "createdAt")
    ActionChains(chrome).click(created_at).perform()

    between = Select(created_at)
    between.select_by_visible_text("between")

    #date1 = chrome.find_element(By.CSS_SELECTOR, "div.react-datepicker__day.react-datepicker__day--0{}".format(yesterday))
    #date2 = chrome.find_element(By.CSS_SELECTOR, "div.react-datepicker__day.react-datepicker__day--0{}".format(friday))

    yesterday_label = 'Choose {}, {} {}{}, {}'.format(yesterday_name, yesterday_month_name, yesterday.lstrip('0'), yesterday_suffix, yesterday_year)
    friday_label = 'Choose {}, {} {}{}, {}'.format(friday_name, friday_month_name, friday.lstrip('0'), friday_suffix, friday_year)

    date1 = chrome.find_element(By.CSS_SELECTOR, "[aria-label='{}']".format(yesterday_label))
    date2 = chrome.find_element(By.CSS_SELECTOR, "[aria-label='{}']".format(friday_label))
    
    if today == "Monday":
        ActionChains(chrome).click(date2).perform()
        ActionChains(chrome).click(date1).perform()
    else:
        ActionChains(chrome).click(date1).perform()
        ActionChains(chrome).click(date1).perform()

    chrome.find_element(By.CSS_SELECTOR, "button.sc-18886227-0.buupqE.button.button--kind-primary.button--size-normal.button--has-label").click()

    wait.until(lambda chrome: chrome.find_element(By.CSS_SELECTOR, "button.sc-4134bb64-0.eGuDRB"))

    chrome.find_element(By.CSS_SELECTOR, "button.sc-4134bb64-0.eGuDRB").click()

    wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"sc-4134bb64-0").is_enabled())

    time.sleep(3)

    print("Report Emailed!")


#Homebase Login function
def homebaseLogin(email, password):
    chrome.get("https://joinhomebase.com/")

    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.button.simple")))

    chrome.find_element(By.CSS_SELECTOR, "a.button.simple").click()

    wait.until(lambda driver: driver.find_element(By.ID, "account_login"))

    account_login = chrome.find_element(By.ID, "account_login")
    account_password = chrome.find_element(By.ID, "account_password")
    account_login.send_keys(email)
    account_password.send_keys(password)
    account_password.send_keys(Keys.ENTER)

    print("Logging in to Homebase...")

    wait.until(lambda driver: driver.find_element(By.ID, "dashboard-location-dashboard"))

    print("Logged in!")


#Function to retrieve Today's Schedule csv file
def getSchedule():
    chrome.get(
    "https://app.joinhomebase.com/schedule_builder#day/shift/{}/{}/{}".format(
    dt.datetime.now().strftime("%m"),  
    dt.datetime.now().strftime("%d"), 
    dt.datetime.now().strftime("%Y")))

    time.sleep(1)

    wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "a.btn.btn-primary-action.schedule-header__button.export-roster"))
    chrome.find_element(By.CSS_SELECTOR, "a.btn.btn-primary-action.schedule-header__button.export-roster").click()

    print("Download Started!")

    wait.until(lambda driver: any(f.startswith('daily-roster') and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))

    print("Download Complete!")


#Function to retrieve Timesheets csv file
def getTimesheets():
    
    if today == "Monday":
        chrome.get("https://app.joinhomebase.com/timesheets?endDate={}{}{}{}{}{}{}{}{}{}{}".format( 
        dt.datetime.now().strftime("%m"), 
        f"%2F", 
        friday, 
        f"%2F", 
        dt.datetime.now().strftime("%Y"), 
        f"&filter=all&groupBy=role&startDate=", 
        dt.datetime.now().strftime("%m"), 
        f"%2F", 
        yesterday, 
        f"%2F", 
        dt.datetime.now().strftime("%Y")))
    else:
        chrome.get("https://app.joinhomebase.com/timesheets?endDate={}{}{}{}{}{}{}{}{}{}{}".format( 
        dt.datetime.now().strftime("%m"), 
        f"%2F", 
        yesterday, 
        f"%2F", 
        dt.datetime.now().strftime("%Y"), 
        f"&filter=all&groupBy=role&startDate=", 
        dt.datetime.now().strftime("%m"), 
        f"%2F", 
        yesterday, 
        f"%2F", 
        dt.datetime.now().strftime("%Y")))

    time.sleep(1)

    wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "button.Button.Button--medium.Button--theme-primary.TimesheetsNavigation__download"))
    try:
        chrome.find_element(By.CSS_SELECTOR, "button.Button.Button--medium.Button--theme-primary.TimesheetsNavigation__download").click()
    except ElementClickInterceptedException:
        time.sleep(2)
        chrome.find_element(By.CSS_SELECTOR, "button.Button.Button--medium.Button--theme-primary.TimesheetsNavigation__download").click()

    wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "button.Button.Button--full-width.Button--medium.Button--theme-primary-purple.DownloadTimesheetsModal__next_button"))
    chrome.find_element(By.XPATH, "//input[@name='detailed']").click()
    chrome.find_element(By.CSS_SELECTOR, "button.Button.Button--full-width.Button--medium.Button--theme-primary-purple.DownloadTimesheetsModal__next_button").click()

    print("Download Started!")

    wait.until(lambda driver: any(f.startswith('El Famoso') and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))

    print("Download Complete!")


#Function to retrieve csv from download link in email
def getShipments(email, password):
    chrome.get("https://mail.google.com/mail/?ui=html")

    wait.until(lambda chrome: chrome.find_element(By.NAME, "identifier"))

    gmail = chrome.find_element(By.NAME, "identifier")
    gmail.send_keys(email)
    time.sleep(1)
    gmail.send_keys(Keys.ENTER)

    wait.until(EC.element_to_be_clickable((By.NAME, "Passwd")))

    passwd = chrome.find_element(By.NAME, "Passwd")
    passwd.send_keys(password)
    passwd.send_keys(Keys.ENTER)
    time.sleep(2)

    tries = 0

    #Recursive function to refresh the page until the element is available.
    def findEmail(find_counter = tries):
        find_counter += 1
        if find_counter > 12:
            print("Email not found within 2 minutes")
            return
        else:
            try:
                print("...")
                chrome.find_element(By.XPATH, "//b[contains(text(),'Export shipments report')]").click()
            except NoSuchElementException:                
                time.sleep(10)
                chrome.refresh()
                findEmail(find_counter)
        return
    
    print("searching for email...")

    findEmail()

    wait.until(lambda chrome: chrome.find_element(By.LINK_TEXT, "View report"))

    print("Email Found!")

    chrome.find_element(By.LINK_TEXT, "View report").click()

    print("Download Started!")

    wait.until(lambda driver: any(f.startswith(dt.datetime.now().strftime("%Y")) and f.endswith('.csv') for f in os.listdir(os.path.join(os.path.expanduser("~"), "Downloads"))))

    print("Download Complete!")

    time.sleep(1)


#Function to distribute downloaded files to correct directories
def moveFiles():
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
        time.sleep(1)

    # distribute the files
    src_path = os.path.join("G:\\Shared drives", "Metric Data")
    for file in os.listdir(src_path):
        if file.startswith("pending"):
            move(os.path.join(src_path, file), os.path.join(src_path, "Pending Shipments"))
            print("Moved {} to {}".format(file, os.path.join(dst_path, "Pending Shipments")))
            time.sleep(1)
        elif file.startswith("product_locations"):
            move(os.path.join(src_path, file), os.path.join(src_path, "Product Locations"))
            print("Moved {} to {}".format(file, os.path.join(dst_path, "Product Locations")))
            time.sleep(1)
        elif file.startswith("product_table"):
            move(os.path.join(src_path, file), os.path.join(src_path, "Current Inventory Table"))
            print("Moved {} to {}".format(file, os.path.join(dst_path, "Current Inventory Table")))
            time.sleep(1)
        elif file.startswith("daily-roster"):
            move(os.path.join(src_path, file), os.path.join(src_path, "Today's Schedule"))
            print("Moved {} to {}".format(file, os.path.join(dst_path, "Today's Schedule")))
            time.sleep(1)
        elif file.startswith("El Famoso"):
            move(os.path.join(src_path, file), os.path.join(src_path, "Timesheet Test"))
            print("Moved {} to {}".format(file, os.path.join(dst_path, "Timesheet Test")))
            time.sleep(1)
        elif file.startswith("shipped_items"):
            move(os.path.join(src_path, file), os.path.join(src_path, "Shipped Items"))
            print("Moved {} to {}".format(file, os.path.join(dst_path, "Shipped Items")))
            time.sleep(1)
        elif file.startswith("20{}".format(dt.datetime.now().strftime("%y"))):
            move(os.path.join(src_path, file), os.path.join(src_path, "Shipment Tables"))
            print("Moved {} to {}".format(file, os.path.join(dst_path, "Shipment Tables")))
            time.sleep(1)
        else:
            time.sleep(1)


#Function to update queries in the root data workbooks
def updateData():
    #Open Pending Shipments
    wb = xl.Workbooks.Open(os.path.join('G:\\Shared drives\\Reports\\Dashboard Reports\\', 'Pending Shipments.xlsx'))
    
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


#Function to update the Dashboard workbook and export to pdf
def updateDashboard():
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


#Function to send the Dashboard to the slack channel
def slackDashboard(slack_access_token, channel_id):
    # define PDF file path
    file_path = os.path.join(
        "G:\\Shared drives\\Dashboard Drive\\Dashboards", 
        "Dashboard {}{}{}{}".format( 
        dt.datetime.now().strftime("%y"), 
        dt.datetime.now().strftime("%m"), 
        dt.datetime.now().strftime("%d"), 
        '.pdf')
    ) 
    print("Sending {} to slack...".format(file_path))
    client = WebClient(token=slack_access_token)
    logger = logging.getLogger(__name__)

    try:
        result = client.files_upload(file=file_path, channels="daily-dashboard")
        print(result)
        logger.info(result)
        print("Slack sent!")
    except SlackApiError as e:
        logger.error("Error uploading file: {}".format(e))
        print("Error uploading file: {}".format(e))


#Function to quit selenium
def quitChrome():
    print("Exiting Chrome...")
    chrome.quit()