import time
from selenium import webdriver
import pyautogui
import re
import tkinter as tk
from tkinter import simpledialog
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def get_user_input(prompt):
    root = tk.Tk()
    root.withdraw()
    user_input = simpledialog.askstring(title="Dane użytkownika", prompt=prompt)
    return user_input

imie = get_user_input("Podaj imię:").upper().strip()
nazwisko = get_user_input("Podaj nazwisko:").upper().strip()

# Google Sheets authentication
SERVICE_ACCOUNT_FILE = 'excel.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1k4UVgLa00Hqa7le3QPbwQMSXwpnYPlvcEQTxXqTEY4U'
SHEET_NAME_1 = 'ZP dane kont'
SHEET_NAME_2 = 'ZP status'

# Authenticate and initialize the Google Sheets client
credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)
client = gspread.authorize(credentials)
ws = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME_1)
ws2 = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME_2)

row = None

data = ws.get_all_values()
for i, row_data in enumerate(data):
    if row_data[0].upper().strip() == imie and row_data[1].upper().strip() == nazwisko:
        row = i + 1
        break
if row is None:
    print("Nie znaleziono danych dla podanej osoby.")

def clean_data(cell_value):
    if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
        return None
    if isinstance(cell_value, str):
        return re.sub(r'[.,\'"]', '', cell_value).strip()
    return cell_value

def clean_data(cell_value):
    if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
        return None
    if isinstance(cell_value, str):
        return re.sub(r'[.,\'"]', '', cell_value).strip()
    return cell_value


if 1==1:

    # Podział numeru konta bankowego
    numer = clean_data(data[row-1][6])
    numer=str(numer)
    part1 = numer[:5]
    print(part1)
    part2 = numer[5:10]
    print(part2)
    part3 = numer[10:15]
    part4 = numer[15:20]
    part5 = numer[20:25]
    part6 = numer[25:30]
    part7 = numer[30:]

    numer = clean_data(data[row-1][10])
    part1 =numer[:2]
    print(part1)

    part2 =numer[3:6]
    print(part2)

    part3 =numer[7:10]
    print(part3)

    part4 =numer[11:]
    print(part4)

