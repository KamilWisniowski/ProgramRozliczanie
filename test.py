import pyautogui
import time

try:
    while True:
        x, y = pyautogui.position()
        print(f'Współrzędne kursora: ({x}, {y})')
        time.sleep(1)
except KeyboardInterrupt:
    print("\nProgram zakończony.")