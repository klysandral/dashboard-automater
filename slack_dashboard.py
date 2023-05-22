import logging
import os

import datetime as dt
from datetime import timedelta

from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

from utils import slackDashboard
from hash import slack_key

slackDashboard(slack_key, "C03U3Q0SKUM")