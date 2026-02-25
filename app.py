import os
from flask import Flask, redirect, request, session, url_for, jsonify
import requests
import msal

from engine import insert_new_row_and_mark

app = Flask(__name__)

# Mandatory for Flask session cookies
app.secret_key = os.environ.get("SECRET_KEY", "change-me")

CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")

BASE_URL = os.environ.get("BASE_URL", "https://personal-agent-zymv.onrender.com").rstrip("/")

TARGET_FILE = os.environ.get("TARGET_FILE", "Master_Dorin_Agent.xlsx")
APP_FILE_NAME = TARGET_FILE

RUN_TOKEN = os.environ.get("RUN_TOKEN", "")

# For personal + work accounts
AUTHORITY = "https://login.microsoftonline.com/common"
REDIRECT_PATH = "/callback"
REDIRECT_URI = BASE_URL + REDIRECT_PATH

# AppFolder
