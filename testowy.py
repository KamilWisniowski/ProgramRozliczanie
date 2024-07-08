import os
import time
import base64
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import simpledialog

def get_user_input(prompt):
    root = tk.Tk()
    root.withdraw()
    user_input = simpledialog.askstring(title="Dane użytkownika", prompt=prompt)
    root.destroy()
    return user_input

def utworz_strukture(imie, nazwisko, rok_rozliczenia):
    sciezka_glowna = os.path.join('C:\\Users\\Kamil\\Desktop\\biuro', f"{imie} {nazwisko}")
    sciezka_rok = os.path.join(sciezka_glowna, rok_rozliczenia)
    sciezka_do_druku = os.path.join(sciezka_rok, 'DO DRUKU')
    os.makedirs(sciezka_do_druku, exist_ok=True)
    return sciezka_do_druku

def zapisz_pdf_bin(imie, nazwisko, folder_docelowy, bin_pdf, nazwa_pliku_pdf):
    sciezka_docelowa = os.path.join(folder_docelowy, f"{imie}_{nazwisko}_{nazwa_pliku_pdf}.pdf")
    if os.path.exists(sciezka_docelowa):
        print("Plik już istnieje w folderze.")
    else:
        try:
            with open(sciezka_docelowa, 'wb') as f:
                f.write(bin_pdf)
            print(f"Plik zapisany jako: {sciezka_docelowa}")
        except Exception as e:
            print(f"Błąd podczas zapisywania pliku: {e}")

def pierwsza_zgoda(driver):
    driver.maximize_window()
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="datenschutz"]/button'))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/div[2]/ul/li[3]/div/div/div[1]/a/div/div[2]'))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="lip_formcatalog"]/div[2]/div[2]/div[2]/ul/li[5]/div/a/span'))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="lip_formcatalog"]/div[2]/div[2]/div[2]/ul/li[6]/div/a/span'))).click()
    except Exception as e:
        print(f"Błąd podczas klikania zgody: {e}")

def PobieranieFormularza(driver, imie, nazwisko, rok_rozliczenia):
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/form/div[6]/div/div/div[1]/div[2]/div[1]/div/div[3]/input'))).click()
        time.sleep(3)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/form/div[7]/div/div[1]/div/div[2]/p/a'))).click()
        time.sleep(1)
        
        # Pobranie zawartości pliku PDF jako binarna
        bin_pdf = driver.execute_script("""
            const xhr = new XMLHttpRequest();
            xhr.open('GET', arguments[0], false);
            xhr.overrideMimeType('text/plain; charset=x-user-defined');
            xhr.send(null);
            if (xhr.status === 200) {
                return xhr.responseText;
            } else {
                return null;
            }
        """, driver.current_url)
        
        if bin_pdf:
            # Konwersja do formatu binarnego
            bin_pdf = bytes(bin_pdf, 'utf-8').decode('unicode_escape').encode('latin-1')
            nazwa_pliku_pdf = "Hauptvordruck_ESt_1_C"
            folder_docelowy = utworz_strukture(imie, nazwisko, rok_rozliczenia)
            zapisz_pdf_bin(imie, nazwisko, folder_docelowy, bin_pdf, nazwa_pliku_pdf)
        else:
            print("Nie udało się pobrać zawartości pliku PDF.")
    except Exception as e:
        print(f"Błąd podczas pobierania formularza: {e}")

def wczytaj_dane_i_wypelnij(driver):
    imie = get_user_input("Podaj imię:")
    nazwisko = get_user_input("Podaj nazwisko:")
    rok_rozliczenia = get_user_input("Podaj rok rozliczenia:")
    
    if imie and nazwisko and rok_rozliczenia:
        try:
            pierwsza_zgoda(driver)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="lip_formcatalog"]/div[2]/div[2]/div[2]/ul/li[3]/div[1]/a/span[1]/span'))).click()

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/form/div[4]/div/div/div/div[4]/input'))).click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="k1"]'))).click()
            time.sleep(2)
            PobieranieFormularza(driver, imie, nazwisko, rok_rozliczenia)
        except Exception as e:
            print(f"Błąd podczas wczytywania danych i wypełniania formularza: {e}")
    else:
        print("Nie podano wszystkich wymaganych danych.")

if __name__ == "__main__":
    driver = webdriver.Chrome()
    url = 'https://www.formulare-bfinv.de/ffw/form/display.do?%24context=2802C5863D1DB0B5962F'
    driver.get(url)
    wczytaj_dane_i_wypelnij(driver)
    driver.quit()
