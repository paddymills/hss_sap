import pyautogui
import time
import tqdm

QTY_TO_PRINT = 18
QTY_PRINTED = 0

for _ in tqdm.tqdm( range(QTY_TO_PRINT) ):
    for _ in range(QTY_PRINTED):
        continue

    pyautogui.click(25, 205)
    pyautogui.click(1176, 160)
    time.sleep(0.5)
    pyautogui.typewrite("ZE22")
    pyautogui.press("enter")
    time.sleep(1.0)

