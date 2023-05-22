
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime as dt
from win32com.client import Dispatch


options = Options()
options.add_experimental_option("detach", True)
options.add_argument('--ignore-certificate-errors')

# Create a new instance of the Chrome driver
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

driver.get('https://app.shiphero.com/dashboard/orders/pending_shipments')

wait.until(lambda driver: driver.find_element(By.NAME,"email"))

#Find the username and password fields
username_field = driver.find_element(By.NAME, "email")
password_field = driver.find_element(By.NAME, "password")

wait.until(EC.element_to_be_clickable((By.NAME, "email")))

#Enter username and password
username_field.send_keys("nate.s.rhodes@gmail.com")
password_field.send_keys("sl!mel0rd")

#Submit the form
clickable = driver.find_element(By.NAME, "submit")
ActionChains(driver).click(clickable).perform()

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn"))
try:
    driver.execute_script("arguments[0].parentNode.removeChild(arguments[0])", driver.find_element(By.ID, "pending_shipments_processing"))
except:
    pass
driver.find_element(By.CLASS_NAME,"btn").click()

time.sleep(1)

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn").is_enabled())

time.sleep(1)

driver.get('https://app.shiphero.com/dashboard/product-locations?location_level=all')

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn"))
wait.until(lambda driver: driver.find_element(By.CLASS_NAME, "paginate_button"))
time.sleep(1)
driver.find_element(By.CLASS_NAME,"btn").click()

time.sleep(2)

wait.until(lambda driver: driver.find_element(By.CLASS_NAME,"btn").is_enabled())
wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')

time.sleep(2)

driver.get('https://app.shiphero.com/dashboard/products?kit=0&build_kit=0&dropship=0&virtual=0')

time.sleep(1)

wait.until(lambda driver: driver.find_element(By.CLASS_NAME, "load-all-button"))
try:
    driver.execute_script("arguments[0].parentNode.removeChild(arguments[0])", driver.find_element(By.CLASS_NAME, "dataTables_length"))
except:
    pass
driver.find_element(By.CLASS_NAME, "load-all-button").click()

wait.until_not(lambda driver: driver.find_element(By.ID, "your-products-processing"))

driver.find_element(By.LINK_TEXT,"Export All Rows").click

'''
driver.get('https://shipping.shiphero.com/shipments-report/')

wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, "a.auth0-lock-social-button.auth0-lock-social-big-button"))

confirm = driver.find_element(By.CSS_SELECTOR, "a.auth0-lock-social-button.auth0-lock-social-big-button")
ActionChains(driver).click(confirm).perform()

driver.wait_until(lambda driver: driver.find_element(By.CLASS_NAME, "sc-18886227-0"))

filters = driver.find_element(By.CLASS_NAME, "sc-18886227-0")
ActionChains(driver).click(filters).perform()

time.sleep(1)

created_at = driver.find_element(By.NAME, "createdAt")
ActionChains(driver).click(created_at).perform()

between = Select(created_at)
between.select_by_visible_text("between")

yesterday = driver.find_element(By.CSS_SELECTOR, "div.react-datepicker__day.react-datepicker__day--019")
ActionChains(driver).click(yesterday).perform()
ActionChains(driver).click(yesterday).perform()

apply = driver.find_element(By.CSS_SELECTOR, "button.sc-18886227-0.buupqE.button.button--kind-primary.button--size-normal.button--has-label")
ActionChains(driver).click(apply).perform()

driver.wait_until(lambda driver: driver.find_element(By.CLASS_NAME, "sc-4134bb64-0"))

send = driver.find_element(By.CLASS_NAME, "sc-4134bb64-0")
ActionChains(driver).click(send).perform()

time.sleep(10)

driver.get("https://www.google.com/gmail/about")

sign_in = driver.find_element(By.CSS_SELECTOR, "a.button.button--medium.button--mobile-before-hero-only")
ActionChains(driver).click(sign_in).perform()

driver.wait_until(lambda driver: driver.find_element(By.NAME, "identifier"))

gmail = driver.find_element(By.NAME, "identifier")
gmail.send_keys("nathan.rhodes@el-famoso.com")
gmail.send_keys(Keys.ENTER)




#Wait for 5 minutes
time.sleep(300)

# Open excel files from google shared drive
xl = Dispatch("Excel.Application")
xl.Visible = True
xl.Workbooks.Open(os.path.join('G:\\Shared drives\\Reports\\Dashboard Reports\\', 'Pending Shipments.xlsx'))

# Wait
time.sleep(61)

# Save the file
xl.ActiveWorkbook.Save()

# Close the file
xl.ActiveWorkbook.Close()

# Quit the application
xl.Application.Quit()

xl = Dispatch("Excel.Application")
xl.Visible = True
xl.Workbooks.Open(os.path.join('G:\\Shared drives\\Reports\\Dashboard Reports\\', 'Last 7.xlsx'))

# Wait
time.sleep(180)

# Save the file
xl.ActiveWorkbook.Save()

# Close the file
xl.ActiveWorkbook.Close()

# Quit the application
xl.Application.Quit()

'''