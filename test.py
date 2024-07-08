import pyautogui
import time

try:
    while True:
        # Pobieramy aktualną pozycję kursora
        # Pobieramy aktualną pozycję kursora
        x, y = pyautogui.position()
        
        # Wypisujemy aktualne współrzędne kursora
        print(f'Pozycja kursora: x={x}, y={y}')
        
        # Czekamy 1 sekundę
        
        # Wypisujemy aktualne współrzędne kursora
        print(f'Pozycja kursora: x={x}, y={y}')
        
        # Czekamy 1 sekundę
        time.sleep(1)
        
        
except KeyboardInterrupt:
    print("\nPrzerwano program.")

    print("\nPrzerwano program.")
