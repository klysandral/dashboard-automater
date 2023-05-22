import time
import os

import datetime as dt
from datetime import timedelta

from shutil import move

from utils import moveFiles

yesterday = (dt.datetime.now() - timedelta(days=1)).strftime("%d")

moveFiles()