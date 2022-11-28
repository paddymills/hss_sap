
import dostuff
import screenshots

import pyautogui
from tqdm import tqdm

img = screenshots.getScreenShotCollection()

MATL_DOC_COORD = (126, 217)

docs = dostuff.read_sort_min_file("matldocs.txt")

with tqdm(docs) as progress:
    for matl_doc in progress:
        progress.write("Cancelling {}".format(matl_doc))

        dostuff.waitUntilPresent(img.MBST.InitialScreen)
        pyautogui.click(*MATL_DOC_COORD)
        pyautogui.typewrite(matl_doc)
        pyautogui.press("enter")
        
        dostuff.waitUntilPresent(img.MBST.SelectionScreen)
        pyautogui.hotkey("ctrl", "s")

