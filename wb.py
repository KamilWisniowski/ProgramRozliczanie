import pyautogui
import time

print("Przesuń mysz na okno dialogowe zapisu. Skrypt zacznie pobierać współrzędne za 5 sekund...")
time.sleep(5)  # Daj czas na przesunięcie myszy na okno dialogowe

while True:
    x, y = pyautogui.position()
    print(f'Współrzędne kursora to: x={x}, y={y}')
    time.sleep(1)  # Odswieżanie co sekundę