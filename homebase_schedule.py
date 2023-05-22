import time
import os

import datetime as dt
from datetime import timedelta

from cryptography.fernet import Fernet

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.support.ui import Select

from utils import homebaseLogin, getSchedule, quitChrome

from hash import key, email2, pass2


#Set chromedriver options
options = Options()
options.add_experimental_option("detach", True)
options.add_argument('--ignore-certificate-errors')
options.add_argument("--start-maximized")


f = Fernet(key)


homebaseLogin(str(f.decrypt(email2))[2:-1], str(f.decrypt(pass2))[2:-1])
getSchedule()
quitChrome()