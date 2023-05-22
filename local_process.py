import os
import logging
import time

import datetime as dt
from datetime import timedelta

import win32com.client
from shutil import move

from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

from utils import moveFiles, updateDashboard, updateData, slackDashboard

yesterday = (dt.datetime.now() - timedelta(days=1)).strftime("%d")
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False

#Distribute data to shared driver
moveFiles()


#Update workbooks and export PDF
updateData()
updateDashboard()

xl.Quit()


#Send dashboard
slackDashboard("xoxb-52563538771-4691190848999-d407QIobcljJZtnLhXKSvTa5", "C03U3Q0SKUM")