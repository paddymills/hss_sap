
import time
import tqdm

import pyautogui
import pytesseract
import numpy
from PIL import Image, ImageGrab

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"

FIRST_LINE_SELECT = (25, 205)
LINE_X = 25
PRINT_BUTTON = (1176, 160)
PRINTER = "ZE22"

class TagPrinter:

    def __init__(self):
        self._print_btn = pyautogui.locateCenterOnScreen("print_button.png")
        self._mat_x, _ = pyautogui.locateCenterOnScreen("material_header.png")

        self._printls = dict()
        with open("printlist.txt", 'r') as f:
            for l in f.readlines():
                p, q = l.split(',')
                if p not in self._printls:
                    self._printls[p] = 0
                self._printls[p] += int(q)

    def main(self):
        reg = [214, 193, 200, 15]
        while True:
            val = p.read_region(reg)
            if not val:
                break

            try:
                self.print_many(self._printls[val], reg[1] + 5, mark=val)
                del self._printls[val]

            except KeyError:
                print(val, "to be skipped")
            
            reg[1] += 19

    def print_one(self):
        pyautogui.click(*self._print_btn)
            
        time.sleep(0.5)
        pyautogui.typewrite(PRINTER)
        pyautogui.press("enter")


    def print_many(self, qty, y, mark=""):
        for _ in tqdm.trange( qty, desc=mark ):
            pyautogui.click(LINE_X, y)
            self.print_one()
            time.sleep(1.0)
            
    def read_region(self, xywh):
        try:
            capture = pyautogui.screenshot(region=xywh)

            return pytesseract.image_to_string(capture).split()[0].replace('.', '')
        except:
            return None

if __name__ == "__main__":
    p = TagPrinter()
    p.main()
