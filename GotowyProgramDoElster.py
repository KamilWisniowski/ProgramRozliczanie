import time
from selenium import webdriver
import pyautogui
from openpyxl import load_workbook
import tkinter as tk
from tkinter import simpledialog
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import re
from oauth2client.service_account import ServiceAccountCredentials
import gspread
def get_user_input(prompt):
    root = tk.Tk()
    root.withdraw()
    user_input = simpledialog.askstring(title="Dane użytkownika", prompt=prompt)
    return user_input

imie = get_user_input("Podaj imię:").upper().strip()
nazwisko = get_user_input("Podaj nazwisko:").upper().strip()
rok_rozliczenia = get_user_input("Podaj rok:").strip()

# Google Sheets authentication
SERVICE_ACCOUNT_FILE = 'excel.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1k4UVgLa00Hqa7le3QPbwQMSXwpnYPlvcEQTxXqTEY4U'
SHEET_NAME_1 = 'ZP dane kont'
SHEET_NAME_2 = 'ZP status'

# Authenticate and initialize the Google Sheets client
credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)
client = gspread.authorize(credentials)
ws1 = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME_1)
ws2 = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME_2)

def wait_and_send_keys(xpath, value):
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    element.send_keys(value)
def clean_data(cell_value):
    if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
        return None
    if isinstance(cell_value, str):
        return re.sub(r'[.,\'"]', '', cell_value).strip()
    return cell_value

data = ws1.get_all_values()
row = None
for i, row_data in enumerate(data):
    if row_data[0] == imie and row_data[1] == nazwisko:
        row = i + 1
        break

if row is None:
    print("Nie znaleziono danych dla podanej osoby.")
    driver.quit()
    exit()

status_data = ws2.get_all_values()
row2 = None
for i, row_data2 in enumerate(status_data):
    full_name = row_data2[0]
    if full_name:
        full_name_parts = full_name.split()
        if full_name_parts[0] == nazwisko and full_name_parts[1] == imie  and row_data2[2] == rok_rozliczenia:
            row2 = i + 1
            break


imie = imie.upper().strip()
nazwisko = nazwisko.upper().strip()
biuro = data[row-1][2].upper().strip()
nr_telefonu = data[row-1][3]
email = data[row-1][4]
stan_cywilny = data[row-1][5].upper().strip()
numer_konta_bankowego = data[row-1][6]
swift = data[row-1][7].upper().strip()
finanzamt = data[row-1][8].upper()
steuernummer = data[row-1][9]
if steuernummer == "NOCH KEINE":
    steuernummer = None
id_nr_meza = data[row-1][10]
id_nr_meza = id_nr_meza.replace(" ", "").upper()
id_nr_zony = data[row-1][11]
id_nr_zony = id_nr_zony.replace(" ", "").upper()
data_urodzenia_meza = data[row-1][12]
religia = data[row-1][13].upper().strip()
ulica = data[row-1][14].upper()
miejscowosc = data[row-1][15].upper()
data_slubu = data[row-1][16]
data_urodzenia_zony = data[row-1][17]
imie_zony = data[row-1][18].upper()

imie_nazwisko_numer = clean_data(status_data[row2-1][0])
status_de = clean_data(status_data[row2-1][1]).upper()
rok_rozliczenia = clean_data(status_data[row2-1][2])
zwrot = clean_data(status_data[row2-1][3])
opiekun = clean_data(status_data[row2-1][4]).upper()
uwagi = clean_data(status_data[row2-1][5])
poinformowany = clean_data(status_data[row2-1][6]).upper().strip()
wyslany = clean_data(status_data[row2-1][7]).upper().strip()
fahrkosten = clean_data(status_data[row2-1][8])
ubernachtung = clean_data(status_data[row2-1][9])
h24 = clean_data(status_data[row2-1][10])
h8 = clean_data(status_data[row2-1][11])
w_kabinie = clean_data(status_data[row2-1][12])
an_und_ab = clean_data(status_data[row2-1][13])
kind = status_data[row2-1][14]
cena = clean_data(status_data[row2-1][15]).strip()
status_platnosci = clean_data(status_data[row2-1][16]).upper().strip()
zaplacono = clean_data(status_data[row2-1][17])
forma_zaplaty = clean_data(status_data[row2-1][18]).upper().strip()
nr_faktury = clean_data(status_data[row2-1][19])
data_wystawienia_faktury = clean_data(status_data[row2-1][20])
ZarobkiMezaNiem = clean_data(status_data[row2-1][21])
ZarobkiMezaNiem = int(ZarobkiMezaNiem)
ZarobkiZonyNiem = clean_data(status_data[row2-1][22])
nr22 = clean_data(status_data[row2-1][23])
nr23 = clean_data(status_data[row2-1][24])
nr25 = clean_data(status_data[row2-1][25])
nr26 = clean_data(status_data[row2-1][26])
nr27 = clean_data(status_data[row2-1][27])
pracodawca = clean_data(status_data[row2-1][28])
chorobowe = clean_data(status_data[row2-1][29])
czy_konto_elster = clean_data(status_data[row2-1][48]).upper().strip()
czy_ogr_ob_podatkowy = clean_data(status_data[row2-1][49]).upper().strip()

klasa_pit1 = clean_data(status_data[row2-1][30])
if klasa_pit1:
    klasa_pit1 = int(klasa_pit1)
else:
    klasa_pit1 = 0
brutto_pit1 = clean_data(status_data[row2-1][31])
if brutto_pit1:
    brutto_pit1 = int(brutto_pit1)
else:
    brutto_pit1 = 0
podatek_pit1 = clean_data(status_data[row2-1][32])
if podatek_pit1:
    podatek_pit1 = int(podatek_pit1)
else:
    podatek_pit1 = 0
doplata_pit1 = clean_data(status_data[row2-1][33])
if doplata_pit1:
    doplata_pit1 = int(doplata_pit1)
else:
    doplata_pit1 = 0
koscielny_pit1 = clean_data(status_data[row2-1][34])
if koscielny_pit1:
    koscielny_pit1 = int(koscielny_pit1)
else:
    koscielny_pit1 = 0
kurzarbeitgeld_pit1 = clean_data(status_data[row2-1][35])
if kurzarbeitgeld_pit1:
    kurzarbeitgeld_pit1 = int(kurzarbeitgeld_pit1)
else:
    kurzarbeitgeld_pit1 = 0

klasa_pit2 = clean_data(status_data[row2-1][36])
if klasa_pit2:
    klasa_pit2 = int(klasa_pit2)
else:
    klasa_pit2 = 0
brutto_pit2 = clean_data(status_data[row2-1][37])
if brutto_pit2:
    brutto_pit2 = int(brutto_pit2)
else:
    brutto_pit2 = 0
podatek_pit2 = clean_data(status_data[row2-1][38])
if podatek_pit2:
    podatek_pit2 = int(podatek_pit2)
else:
    podatek_pit2 = 0
doplata_pit2 = clean_data(status_data[row2-1][39])
if doplata_pit2:
    doplata_pit2 = int(doplata_pit2)
else:
    doplata_pit2 = 0
koscielny_pit2 = clean_data(status_data[row2-1][40])
if koscielny_pit2:
    koscielny_pit2 = int(koscielny_pit2)
else:
    koscielny_pit2 = 0
kurzarbeitgeld_pit2 = clean_data(status_data[row2-1][41])
if kurzarbeitgeld_pit2:
    kurzarbeitgeld_pit2 = int(kurzarbeitgeld_pit2)
else:
    kurzarbeitgeld_pit2 = 0

klasa_pit3 = clean_data(status_data[row2-1][42])
if klasa_pit3:
    klasa_pit3 = int(klasa_pit3)
else:
    klasa_pit3 = 0
brutto_pit3 = clean_data(status_data[row2-1][43])
if brutto_pit3:
    brutto_pit3 = int(brutto_pit3)
else:
    brutto_pit3 = 0
podatek_pit3 = clean_data(status_data[row2-1][44])
if podatek_pit3:
    podatek_pit3 = int(podatek_pit3)
else:
    podatek_pit3 = 0
doplata_pit3 = clean_data(status_data[row2-1][45])
if doplata_pit3:
    doplata_pit3 = int(doplata_pit3)
else:
    doplata_pit3 = 0
koscielny_pit3 = clean_data(status_data[row2-1][46])
if koscielny_pit3:
    koscielny_pit3 = int(koscielny_pit3)
else:
    koscielny_pit3 = 0
kurzarbeitgeld_pit3 = clean_data(status_data[row2-1][47])
if kurzarbeitgeld_pit3:
    kurzarbeitgeld_pit3 = int(kurzarbeitgeld_pit3)
else:
    kurzarbeitgeld_pit3 = 0
czy_konto_elster = clean_data(status_data[row2-1][48])

def parse_address(address):
    pattern = re.compile(r'(?P<street>[\w\sąćęłńóśźżĄĆĘŁŃÓŚŹŻ]+?)(?:\s+(?P<block>\d+))?(?:\s*(?P<part>[A-Z]))?(?:/(?P<apartment>\d+))?$')
    match = pattern.match(address.strip())
    if match:
        return match.groupdict()
    else:
        raise ValueError("Format adresu jest niepoprawny")
def przekształć_na_nazwisko_dla_żony(nazwisko_męża):
    if not nazwisko_męża:
        return ""

    nazwisko_męża = nazwisko_męża.strip().upper()

    if nazwisko_męża.endswith("SKI"):
        nazwisko_żony = nazwisko_męża[:-3] + "SKA"
    elif nazwisko_męża.endswith("CKI"):
        nazwisko_żony = nazwisko_męża[:-3] + "CKA"
    elif nazwisko_męża.endswith("DZKI"):
        nazwisko_żony = nazwisko_męża[:-4] + "DZKA"
    elif nazwisko_męża.endswith("DZKA"):
        nazwisko_żony = nazwisko_męża
    elif nazwisko_męża.endswith("CKI"):
        nazwisko_żony = nazwisko_męża[:-3] + "CKA"
    elif nazwisko_męża.endswith("DZKA"):
        nazwisko_żony = nazwisko_męża[:-4] + "DZKA"
    elif nazwisko_męża.endswith("EWSKI"):
        nazwisko_żony = nazwisko_męża[:-5] + "EWSKA"
    elif nazwisko_męża.endswith("EWICZ"):
        nazwisko_żony = nazwisko_męża[:-5] + "EWICZ"
    elif nazwisko_męża.endswith("ICZ"):
        nazwisko_żony = nazwisko_męża[:-2] + "ÓWNA"
    elif nazwisko_męża.endswith("Y"):
        nazwisko_żony = nazwisko_męża[:-1] + "A"
    else:
        nazwisko_żony = nazwisko_męża

    return nazwisko_żony.upper()

#============================================================================#
# Włączenie strony do rozliczania (Logowanie)
driver = webdriver.Chrome()
url = 'https://www.elster.de/eportal/login/softpse'
driver.get(url)
driver.maximize_window()

# Wybór metody logowania
driver.find_element(By.XPATH, '//*[@id="submitButton_loginBox.file_cert"]').click()
time.sleep(.5)
pyautogui.click(x=115, y=488)
pyautogui.click(x=412, y=195)
pyautogui.press('enter')

# Wprowadzenie hasła i zatwierdzenie
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys('GrandaKamil11071989!')
driver.find_element(By.XPATH, '//*[@id="bestaetigenButton"]').click()
try:
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="weiterButton"]'))).click()
except TimeoutException:
    pass
try:
    element = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="temporaereaufgaben_nein_button"]'))
    )
    element.click()
except TimeoutException:
    pass

if czy_ogr_ob_podatkowy == "NIE" and rok_rozliczenia=='2023':
    est1aURL= 'https://www.elster.de/eportal/interpreter/fruehereAbgaben/est-2023'
    driver.get(est1aURL)
elif czy_ogr_ob_podatkowy == "TAK" and rok_rozliczenia=='2023':
    est1cURL= 'https://www.elster.de/eportal/interpreter/fruehereAbgaben/est12-2023'
    driver.get(est1cURL)
else:
    print("Rok rozliczenia nie jest obsługiwany")
    exit()

#POMINIĘCIĘ SCIĄGANIA DANYCH
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Continue"]'))).click()

try:
    WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="temporaereaufgaben_nein_button"]'))).click()
except TimeoutException:
    pass
try:
    WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ExitWizard"]'))).click()
except TimeoutException:
    pass

#ZAZNACZANIE POTRZEBNYCH FORMULARZY
if kind is not None and czy_ogr_ob_podatkowy == "NIE":
    checkbox_kind = driver.find_element(By.XPATH, '//*[@id="VAnlageKind"]')
    if not checkbox_kind.is_selected():
        # Zaznaczamy checkbox
        checkbox_kind.click()
if czy_ogr_ob_podatkowy == "NIE":
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="VAnlageSA"]'))).click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="VAnlageWAESt"]'))).click()
    pyautogui.click(x=1910, y=501)
    pyautogui.click(x=826, y=428)
elif czy_ogr_ob_podatkowy == "TAK":
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="VAnlageN"]'))).click()
    
checkbox_Vorsorge = driver.find_element(By.XPATH, '//*[@id="VAnlageVor"]')
if not checkbox_Vorsorge.is_selected():
    # Zaznaczamy checkbox
    checkbox_Vorsorge.click()

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Continue"]'))).click()
try:
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Cancel"]'))).click()
except TimeoutException:
    pass

#chyba do usunięcia
time.sleep(.5)
pyautogui.click(x=1906,y=518)
time.sleep(1)
pyautogui.click(x=315,y=333)
time.sleep(1)
try:
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="fillInProfile_Startseite-MeinProfil"]'))).click()
except:
    pass
time.sleep(1)
checkbox_Einkommensteuererklärung = driver.find_element(By.XPATH, '//*[@id="Startseite(0)_fields(eruESt1AArt_ErklE0100001)"]')
if not checkbox_Einkommensteuererklärung.is_selected():
    checkbox_Einkommensteuererklärung.click()

#WYEPŁACZANIE DANYCH dla EST1A
if czy_ogr_ob_podatkowy=="NIE":
    daneURL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VHauptvordruck/StpflPerson'
    driver.get(daneURL)
    time.sleep(1)
    if religia =="VD":
        wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_StpflPerson(0)_fields(eruESt1AAllgAE0100402)"]', 'nicht')
    elif religia =="RK":
        wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_StpflPerson(0)_fields(eruESt1AAllgAE0100402)"]', 'r')
    if data_slubu is not None:
        wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_StpflPerson(0)_fields(eruESt1AAllgAE0100701)"]', data_slubu)
        pyautogui.press('enter')
        time.sleep(1)
    if stan_cywilny=="ŻONATY":
        zonaURL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VHauptvordruck/Lebenspartner'
        driver.get(zonaURL)
        time.sleep(1)

        wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Lebenspartner(0)_fields(eruESt1AAllgBE0100082)"]', id_nr_zony)
        wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Lebenspartner(0)_fields(eruESt1AAllgBE0101001)"]', data_urodzenia_zony)
        nazwisko_zony = przekształć_na_nazwisko_dla_żony(nazwisko)
        wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Lebenspartner(0)_fields(eruESt1AAllgBE0100901)"]', nazwisko_zony)
        wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Lebenspartner(0)_fields(eruESt1AAllgBE0100801)"]', imie_zony)
        if religia=="VD":
            wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Lebenspartner(0)_fields(eruESt1AAllgBE0101002)"]', 'nicht')
        elif religia=="RK":
            wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Lebenspartner(0)_fields(eruESt1AAllgBE0101002)"]', 'r')
        pyautogui.press('enter')
        time.sleep(.3)
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="NextPage"]'))).click()
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VHauptvordruck(0)_Veranlagungsart(0)_fields(eruESt1AAllgVlg_ArtE0101201)"]'))).click()
        pyautogui.press('enter')
        time.sleep(1)
    stronaKontaBankowegoURL='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VHauptvordruck/Bankverbindung'
    driver.get(stronaKontaBankowegoURL)
    wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Bankverbindung(0)_fields(eruESt1AAllgBVE0102603)"]', numer_konta_bankowego)
    wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Bankverbindung(0)_fields(eruESt1AAllgBVE0102201)"]', swift)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VHauptvordruck(0)_Bankverbindung(0)_fields(eruESt1AAllgBVKto_InhE0101601)"]'))).click()
    time.sleep(.1)
    pyautogui.press('enter')
    time.sleep(2)
elif czy_ogr_ob_podatkowy=="TAK":
    daneURL = 'https://www.elster.de/eportal/interpreter/eingabe/est12-2023/Startseite/VHauptvordruck'


if chorobowe:
    time.sleep(1)
    choroboweURL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VHauptvordruck/Einkommensersatzleistungen'
    driver.get(choroboweURL)
    time.sleep(1)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="edit_btn_Startseite(0)_VHauptvordruck(0)_Einkommensersatzleistungen(0)_MZBEinkErs(0)"]'))).click()
    wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Einkommensersatzleistungen(0)_MZBEinkErs(0)_EinkErsInl(0)_fields(eruESt1AEink_ErsInlEinzE0104110)"]', "Krankengeld")
    wait_and_send_keys('//*[@id="Startseite(0)_VHauptvordruck(0)_Einkommensersatzleistungen(0)_MZBEinkErs(0)_EinkErsInl(0)_fields(eruESt1AEink_ErsInlEinzE0104113)"]', chorobowe)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CreateMzbItem/Startseite[0]/VHauptvordruck[0]/Einkommensersatzleistungen[0]/MZBEinkErs[0]/EinkErsInl[0]"]'))).click()

#-------------------------------------------------------------------------------------#
# 045 KIND 
if kind is not None:
    kindURL= 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageKind'
    driver.get(kindURL)
    time.sleep(1)
    try:
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="uebernehmenButton_451563490"]'))).click()
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Cancel"]'))).click()
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="closeButton_datenuebernahmeModal"]'))).click()
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="anlagenAuswahlVertical"]'))).click()
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="VAnlageKind"]'))).click()
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="VAnlageSA"]'))).click()
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Continue"]'))).click()
    except:
        pass
    
    def process_sentence(sentence):
        people = []
        # Podziel tekst według nowej linii
        lines = sentence.strip().split('\n')
        for line in lines:
            # Usuń białe znaki z początku i końca linii
            line = line.strip()
            if line:
                # Podziel każdą linię według spacji
                parts = line.split()
                if len(parts) >= 4:
                    # Przypisz odpowiednie fragmenty do zmiennych
                    name = parts[0]
                    date = parts[1]
                    status = ' '.join(parts[2:])
                    # Dodaj dane osoby do listy
                    people.append((name, date, status))
                else:
                    print(f"Warning: line '{line}' does not contain enough parts")
        return people, len(people)

    def remove_commas(text):
        return text.replace(",", "")


    def wypelnij_formularz_dla_dziecka(index, person, stan_cywilny):
        driver.get(kindURL)
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, f'//*[@id="JumpToPage/Startseite[0]/MAVSAnlageKind[0]/VAnlageKind[{index-1}]"]'))).click()
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="AngabenZumKind"]'))).click()
        time.sleep(1)
        wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_AngabenZumKind(0)_fields(eruKindAng_KindAllgE0500107)"]', person[0])
        wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_AngabenZumKind(0)_fields(eruKindAng_KindAllgE0500701)"]', person[1])
        if person[2] != 'BEZ KG':
            wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_AngabenZumKind(0)_fields(eruKindAng_KindAllgE0500702)"]', person[2])
        wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_AngabenZumKind(0)_fields(eruKindAng_KindAllgE0500706)"]', 'Sachsen')
        wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_AngabenZumKind(0)_WohnortAusland(0)_fields(eruKindAng_KindWSAuslE0500704)"]', '01.01-31.12')
        wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_AngabenZumKind(0)_WohnortAusland(0)_fields(eruKindAng_KindWSAuslE0500104)"]', 'POLEN')
        pyautogui.press('enter')
        time.sleep(5)
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="NextPage"]'))).click()
        wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_0633dfe3-9fb4-4baa-b163-223d550bfd06(0)_fields(eruKindK_VerhK_Verh_AE0500807)"]', 'lei')
        wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_0633dfe3-9fb4-4baa-b163-223d550bfd06(0)_fields(eruKindK_VerhK_Verh_AE0500601)"]', '01.01-31.12')
        if stan_cywilny == "ŻONATY":
            wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_0633dfe3-9fb4-4baa-b163-223d550bfd06(0)_fields(eruKindK_VerhK_Verh_BE0500808)"]', 'lei')
            wait_and_send_keys(f'//*[@id="Startseite(0)_MAVSAnlageKind(0)_VAnlageKind({index-1})_0633dfe3-9fb4-4baa-b163-223d550bfd06(0)_fields(eruKindK_VerhK_Verh_BE0500805)"]', '01.01-31.12')
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '//*[@id="fbc_FormData://est-2023-v1/Startseite[0]/MAVSAnlageKind[0]"]/span[2]'))).click()

    if len(kind) > 8:
        kind = remove_commas(kind)
        people_data, number_of_people = process_sentence(kind)
        print(people_data)
        print(f"Number of children: {number_of_people}")  # Debugging
        for i, person in enumerate(people_data):
            print(f"Filling form for child {i + 1}: {person}")  # Debugging
            if i >= 5:
                break
            wypelnij_formularz_dla_dziecka(i + 1, person, stan_cywilny)
#-------------------------------------------------------------------------------------#
# 055 Anlage N
# Wprowadzenie danych
# Przejście do sekcji Anlage N

AnlageN = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/AngabenZumArbeitslohn'
driver.get(AnlageN)


sum_brutto_1_5 = 0
sum_podatek_1_5 = 0
sum_doplata_1_5 = 0
sum_koscielny_1_5 = 0
sum_kurzarbeitgeld_1_5 = 0

sum_brutto_6 = 0
sum_podatek_6 = 0
sum_doplata_6 = 0
sum_koscielny_6 = 0
sum_kurzarbeitgeld_6 = 0

# Sumowanie wartości w ramach tej samej grupy klas
if klasa_pit1 in range(1, 6):
    sum_brutto_1_5 += brutto_pit1
    sum_podatek_1_5 += podatek_pit1
    sum_doplata_1_5 += doplata_pit1
    sum_koscielny_1_5 += koscielny_pit1
    sum_kurzarbeitgeld_1_5 += kurzarbeitgeld_pit1
elif klasa_pit1 == 6:
    sum_brutto_6 += brutto_pit1
    sum_podatek_6 += podatek_pit1
    sum_doplata_6 += doplata_pit1
    sum_koscielny_6 += koscielny_pit1
    sum_kurzarbeitgeld_6 += kurzarbeitgeld_pit1

if klasa_pit2 in range(1, 6):
    sum_brutto_1_5 += brutto_pit2
    sum_podatek_1_5 += podatek_pit2
    sum_doplata_1_5 += doplata_pit2
    sum_koscielny_1_5 += koscielny_pit2
    sum_kurzarbeitgeld_1_5 += kurzarbeitgeld_pit2
elif klasa_pit2 == 6:
    sum_brutto_6 += brutto_pit2
    sum_podatek_6 += podatek_pit2
    sum_doplata_6 += doplata_pit2
    sum_koscielny_6 += koscielny_pit2
    sum_kurzarbeitgeld_6 += kurzarbeitgeld_pit2

if klasa_pit3 in range(1, 6):
    sum_brutto_1_5 += brutto_pit3
    sum_podatek_1_5 += podatek_pit3
    sum_doplata_1_5 += doplata_pit3
    sum_koscielny_1_5 += koscielny_pit3
    sum_kurzarbeitgeld_1_5 += kurzarbeitgeld_pit3
elif klasa_pit3 == 6:
    sum_brutto_6 += brutto_pit3
    sum_podatek_6 += podatek_pit3
    sum_doplata_6 += doplata_pit3
    sum_koscielny_6 += koscielny_pit3
    sum_kurzarbeitgeld_6 += kurzarbeitgeld_pit3

if klasa_pit1 is not None:
    time.sleep(1)
    AnlageNURL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/AngabenZumArbeitslohn'
    driver.get(AnlageNURL)
    time.sleep(3)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="edit_btn_Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)"]'))).click()
    time.sleep(3)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="AngabenZumArbeitslohn"]'))).click()

 # Wypełnianie danych ANLAGE N
if klasa_pit1 in range(1, 6) or klasa_pit2 in range(1, 6) or klasa_pit3 in range(1, 6):
    values = [x for x in [klasa_pit1, klasa_pit2, klasa_pit3] if x > 0]
    if values:
        min_value = min(values)
    klasa_pit1=str(klasa_pit1)
    driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_AngabenZumArbeitslohn(0)_fields(eruNArbLLStB_1_5_SumE0200002)"]').send_keys(klasa_pit1)
    
    if sum_brutto_1_5 is not None:
        driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_AngabenZumArbeitslohn(0)_LStBKlassenEinsBisFuenf(0)_fields(eruNArbLLStB_1_5_EinzE0200204)"]').send_keys(sum_brutto_1_5)
    if sum_podatek_1_5 is not None:
        driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_AngabenZumArbeitslohn(0)_LStBKlassenEinsBisFuenf(0)_fields(eruNArbLLStB_1_5_EinzE0200304)"]').send_keys(sum_podatek_1_5)
    if sum_doplata_1_5 is not None:
        driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_AngabenZumArbeitslohn(0)_LStBKlassenEinsBisFuenf(0)_fields(eruNArbLLStB_1_5_EinzE0200404)"]').send_keys(sum_doplata_1_5)
    if sum_koscielny_1_5 is not None:
        driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_AngabenZumArbeitslohn(0)_LStBKlassenEinsBisFuenf(0)_fields(eruNArbLLStB_1_5_EinzE0200504)"]').send_keys(sum_koscielny_1_5)
    elif sum_koscielny_1_5==0:
        pass
    time.sleep(.5)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CreateMzbItem/Startseite[0]/MAVSAnlageN[0]/VAnlageN[0]/AngabenZumArbeitslohn[0]/LStBKlassenEinsBisFuenf[0]"]'))).click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="NextPage"]'))).click()
    time.sleep(1)
if klasa_pit1 == 6 or klasa_pit2 == 6 or klasa_pit3 == 6:
    if sum_brutto_6 is not None:
        driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_AngabenZumArbeitslohn(0)_LStBKlasseSechsUrlaubskasse(0)_fields(eruNArbLLStB_6_EinzE0200202)"]').send_keys(sum_brutto_6)
    if sum_podatek_6 is not None:
        driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0 )_VAnlageN(0)_AngabenZumArbeitslohn(0)_LStBKlasseSechsUrlaubskasse(0)_fields(eruNArbLLStB_6_EinzE0200302)"]').send_keys(sum_podatek_6)
    if sum_doplata_6 is not None:
        driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_AngabenZumArbeitslohn(0)_LStBKlasseSechsUrlaubskasse(0)_fields(eruNArbLLStB_6_EinzE0200402)"]').send_keys(sum_doplata_6)
    if sum_koscielny_6 is not None:
        driver.find_element("xpath", '//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_AngabenZumArbeitslohn(0)_LStBKlasseSechsUrlaubskasse(0)_fields(eruNArbLLStB_6_EinzE0200502)"]').send_keys(sum_koscielny_6)
    elif sum_koscielny_6==0:
        pass
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CreateMzbItem/Startseite[0]/MAVSAnlageN[0]/VAnlageN[0]/AngabenZumArbeitslohn[0]/LStBKlasseSechsUrlaubskasse[0]"]'))).click()
    
sum_kurzarbeitgeld = sum_kurzarbeitgeld_6 + sum_kurzarbeitgeld_1_5
if sum_kurzarbeitgeld > 0:
    kurzarbeitgeldURL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/LohnEntgeltersatzleistungen'
    driver.get(kurzarbeitgeldURL)
    time.sleep(2)
    
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_LohnEntgeltersatzleistungen(0)_Lohnersatzleistungen(0)_fields(eruNArbLErsatzleistBetragEinzE0202605)"]', sum_kurzarbeitgeld)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_LohnEntgeltersatzleistungen(0)_Lohnersatzleistungen(0)_fields(eruNArbLErsatzleistBetragEinzE0202604)"]', 'Kurzarbeitergeld')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CreateMzbItem/Startseite[0]/MAVSAnlageN[0]/VAnlageN[0]/LohnEntgeltersatzleistungen[0]/Lohnersatzleistungen[0]"]'))).click()
    time.sleep(2)
if 1==1:
    ubraniaURL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/Arbeitsmittel'
    driver.get(ubraniaURL)
    time.sleep(2)
    driver.get(ubraniaURL)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_Arbeitsmittel(0)_AufwendungenArbeitsmittel(0)_fields(eruNWkArbeitsmittelEinzE0204401)"]', 'typische Berufskleidung und Reinigungskosten')
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_Arbeitsmittel(0)_AufwendungenArbeitsmittel(0)_fields(eruNWkArbeitsmittelEinzE0204402)"]', '110')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CreateMzbItem/Startseite[0]/MAVSAnlageN[0]/VAnlageN[0]/Arbeitsmittel[0]/AufwendungenArbeitsmittel[0]"]'))).click()
    time.sleep(1)

if fahrkosten is not None:
    fahrkostenUrl='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/ReisekostenAuswaertstaetigkeit'
    driver.get(fahrkostenUrl)
    time.sleep(4)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_ReisekostenAuswaertstaetigkeit(0)_AwtFahrtkosten(0)_fields(eruNWkAWTFahrtE0205003)"]', 'Fahrkosten')
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_ReisekostenAuswaertstaetigkeit(0)_AwtFahrtkosten(0)_fields(eruNWkAWTFahrtE0205004)"]', fahrkosten)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CreateMzbItem/Startseite[0]/MAVSAnlageN[0]/VAnlageN[0]/ReisekostenAuswaertstaetigkeit[0]/AwtFahrtkosten[0]"]'))).click()

if ubernachtung is not None:
    ubernachtungURL='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/ReisekostenAuswaertstaetigkeit'
    driver.get(ubernachtungURL)
    time.sleep(2)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_ReisekostenAuswaertstaetigkeit(0)_AwtUebernachtungskosten(0)_fields(eruNWkAWTUebernachtE0206301)"]', 'Ubernachtungskosten')
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_ReisekostenAuswaertstaetigkeit(0)_AwtUebernachtungskosten(0)_fields(eruNWkAWTUebernachtE0206302)"]', ubernachtung)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CreateMzbItem/Startseite[0]/MAVSAnlageN[0]/VAnlageN[0]/ReisekostenAuswaertstaetigkeit[0]/AwtUebernachtungskosten[0]"]'))).click()

if w_kabinie is not None:
    wKabinieURL='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/ReisekostenAuswaertstaetigkeit'
    driver.get(wKabinieURL)
    time.sleep(4)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_ReisekostenAuswaertstaetigkeit(0)_fields(eruNWkAWTPB_KraftfE0206501)"]', w_kabinie)
    pyautogui.press('enter')

if h24 is not None:
    h24URL='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/PauschbetraegeVerpflegung'
    driver.get(h24URL)
    time.sleep(4)
    driver.get(h24URL)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_PauschbetraegeVerpflegung(0)_fields(eruNWkVMAInlE0205409)"]', h24)
    pyautogui.press('enter')

if h8 is not None:
    h8URL='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/PauschbetraegeVerpflegung'
    driver.get(h8URL)
    time.sleep(4)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_PauschbetraegeVerpflegung(0)_fields(eruNWkVMAInlE0205201)"]', h8)
    pyautogui.press('enter')

if an_und_ab is not None:
    abUndAbURL='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/PauschbetraegeVerpflegung'
    driver.get(abUndAbURL)
    time.sleep(4)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_PauschbetraegeVerpflegung(0)_fields(eruNWkVMAInlE0205302)"]', an_und_ab)
    pyautogui.press('enter')

if pracodawca is not None:
    time.sleep(1)
    pracodawcaURL='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/MAVSAnlageN/VAnlageN/0/PauschbetraegeVerpflegung'
    driver.get(pracodawcaURL)
    time.sleep(3)
    wait_and_send_keys('//*[@id="Startseite(0)_MAVSAnlageN(0)_VAnlageN(0)_PauschbetraegeVerpflegung(0)_fields(eruNWkVMAVMA_ErsatzE0205108)"]', pracodawca)
    time.sleep(1)
    pyautogui.press('enter')
time.sleep(2)

#Kirchensteuer
if koscielny_pit1 or koscielny_pit2 or koscielny_pit3:
    koscielny_pit1 = round(int(koscielny_pit1)) if koscielny_pit1 else 0
    koscielny_pit2 = round(int(koscielny_pit2)) if koscielny_pit2 else 0
    koscielny_pit3 = round(int(koscielny_pit3)) if koscielny_pit3 else 0

    koscielnyURL='https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VAnlageSA/Kirchensteuer'
    driver.get(koscielnyURL)
    koscielny_Caly = koscielny_pit1 + koscielny_pit2 + koscielny_pit3
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageSA(0)_Kirchensteuer(0)_KiSt(0)_fields(eruSAKiStGezahltEinzE0108003)"]', 'Kirchensteuer laut Lohnsteuerbescheinigung steuerpflichtige Person / Ehemann / Person A')
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageSA(0)_Kirchensteuer(0)_KiSt(0)_fields(eruSAKiStGezahltEinzE0108004)"]', koscielny_Caly)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CreateMzbItem/Startseite[0]/VAnlageSA[0]/Kirchensteuer[0]/KiSt[0]"]'))).click()
    time.sleep(2)

#Anlage Vorsorgeaufwand: Angaben zu Vorsorgeaufwendungen
if nr22 is not None or nr23 is not None:
    nr22i23URL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VAnlageVor/BeitraegeZurAltersvorsorge'
    driver.get(nr22i23URL)
    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="delete_btn_readMode_Startseite(0)_VAnlageVor(0)_BeitraegeZurAltersvorsorge(0)_BeitraegeZurAltersvorsorgeMZB(0)"]'))).click()
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="confirm_mzbLoeschenModal"]'))).click()
    except:
        pass
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="edit_btn_Startseite(0)_VAnlageVor(0)_BeitraegeZurAltersvorsorge(0)_BeitraegeZurAltersvorsorgeMZB(0)"]'))).click()
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageVor(0)_BeitraegeZurAltersvorsorge(0)_BeitraegeZurAltersvorsorgeMZB(0)_fields(eruVORAVorE2000401)"]', nr23)
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageVor(0)_BeitraegeZurAltersvorsorge(0)_BeitraegeZurAltersvorsorgeMZB(0)_fields(eruVORAVorE2000801)"]', nr22)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="UpdateMzbItem/Startseite[0]/VAnlageVor[0]/BeitraegeZurAltersvorsorge[0]/BeitraegeZurAltersvorsorgeMZB[0]"]'))).click()

if nr25 is not None:
    nr25i26URL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VAnlageVor/BeitraegeInlGesKrankenPflegevers'
    driver.get(nr25i26URL)
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="delete_btn_Startseite(0)_VAnlageVor(0)_BeitraegeInlGesKrankenPflegevers(0)_MZBBeitraegeInlGesKrankenPflegeversMZB(0)"]'))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="confirm_mzbLoeschenModal"]'))).click()
    except:
        pass
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="edit_btn_Startseite(0)_VAnlageVor(0)_BeitraegeInlGesKrankenPflegevers(0)_MZBBeitraegeInlGesKrankenPflegeversMZB(0)"]'))).click()
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageVor(0)_BeitraegeInlGesKrankenPflegevers(0)_MZBBeitraegeInlGesKrankenPflegeversMZB(0)_fields(eruVORBeitr_g_KV_PV_InlANE2001203)"]', nr25)
    if nr26 is None:
        nr26=0
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageVor(0)_BeitraegeInlGesKrankenPflegevers(0)_MZBBeitraegeInlGesKrankenPflegeversMZB(0)_fields(eruVORBeitr_g_KV_PV_InlANE2001505)"]', nr26)
    pyautogui.press('enter')

if nr27 is not None:
    nr27URL = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VAnlageVor/WeitereSonstigeVorsorgeaufwendungen'
    driver.get(nr27URL)
    try:
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/form/div[5]/main/div/div[4]/div/div/div[1]/div/div[2]/div[2]/div/div/div[3]/div/input[2]'))).click()
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
    except:
        pass
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageVor(0)_WeitereSonstigeVorsorgeaufwendungen(0)_WeitSonsVorAW(0)_fields(eruVORWeit_Sons_VorAWPersE2004403)"]', nr27)
    pyautogui.press('enter')
    time.sleep(2)

#ZAROBKI W POLSCE WA-EST

WaEstUrl = 'https://www.elster.de/eportal/interpreter/eingabe/est-2023/Startseite/VAnlageWAESt'
time.sleep(1)
driver.get(WaEstUrl)
if ZarobkiMezaNiem is None:
        ZarobkiMezaNiem = 0
checkbox1 = driver.find_element(By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0109906)"]')
checkbox2 = driver.find_element(By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStEheg_EU_EWRE0105901)"]')
if ZarobkiZonyNiem is not None:
    ZarobkiZonyNiem = int(ZarobkiZonyNiem)
if ZarobkiZonyNiem is not None:
    if ZarobkiMezaNiem + ZarobkiZonyNiem < 20000:
        if checkbox1.is_selected():
            if checkbox2.is_selected():
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStEheg_EU_EWRE0105901)"]'))).click()
        else:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0109906)"]'))).click()
    else:
        if checkbox2.is_selected():
            if checkbox1.is_selected():
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0109906)"]'))).click()
        else:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStEheg_EU_EWRE0105901)"]'))).click()
else:
    if ZarobkiMezaNiem < 20000:
        if checkbox1.is_selected():
            if checkbox2.is_selected():
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStEheg_EU_EWRE0105901)"]'))).click()
        else:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0109906)"]'))).click()
    else:
        if checkbox2.is_selected():
            if checkbox1.is_selected():
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0109906)"]'))).click()
        else:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStEheg_EU_EWRE0105901)"]'))).click()

try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0105303)"]'))).clear()
except:
    pass
time.sleep(1)
input_field = driver.find_element(By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0105301)"]')
input_field.clear()
time.sleep(.5)
wait_and_send_keys('//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0110306)"]', 'Polen')
wait_and_send_keys('//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0105301)"]', ZarobkiMezaNiem)
if ZarobkiZonyNiem is not None:
    ZarobkiZonyNiem = int(ZarobkiZonyNiem)
    input_field = driver.find_element(By.XPATH, '//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0105303)"]')
    input_field.clear()
    time.sleep(.5)
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0105303)"]', ZarobkiZonyNiem)
    wait_and_send_keys('//*[@id="Startseite(0)_VAnlageWAESt(0)_fields(eruWA_EStAntrag_unb_StpflE0110307)"]', 'Polen')

time.sleep(1)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="SwitchModusPruefen"]'))).click()
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="steuerberechnungLink"]'))).click()
except:
    print("Błąd podczas pobrania strony Steuerberechnung")
input("Wciśnij Enter, aby zakończyć działanie programu...")