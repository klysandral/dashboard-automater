import time
import os

import datetime as dt
from datetime import timedelta

import win32com.client
from shutil import move

from utils import updateDashboard, updateData

yesterday = (dt.datetime.now() - timedelta(days=1)).strftime("%d")
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False

updateData()
updateDashboard()

xl.Quit()
