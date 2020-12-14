import pyautogui
import time
import os
import sys

import uuid
from types import SimpleNamespace
from concurrent import futures

from PIL import Image, ImageGrab
import pytesseract
import numpy

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"

INBOX_HEADER_REGION = (22, 80, 341, 18)
LINE_ITEM_REGION = (352, 233, 102, 12)
WORKFLOW_LOG_ICON_REGION = (524, 156, 15, 14)
WORKFLOW_LOG_HEADER_REGION = (23, 83, 374, 21)

LINE_ITEM_CLICK = (626, 239)
WORKFLOW_LOG_CLICK = (529, 162)
CONTAINER_TAB_CLICK = (498, 600)
EXECUTE_CLICK = (620, 162)

# Popup Get Values
EXPAND_INPUT_DATA_CLICK = (36, 835)
MATL_VALUE_REGION = (205, 850, 150, 18)
WBS_VALUE_REGION = (205, 865, 120, 18)
QTY_VALUE_REGION = (205, 920, 50, 18)

# Run MRP ...
# EXPAND_INPUT_DATA_CLICK = (36, 835)
# MATL_VALUE_REGION = (205, 840, 150, 15)
# WBS_VALUE_REGION = (205, 855, 120, 15)
# QTY_VALUE_REGION = (205, 910, 50, 15)

FAILURE_STRING = "D-1180029-10377"

whileLoopRun = True


def main():
    for f in os.listdir("temp"):
        os.remove("temp\\{}".format(f))

    try:
        getData()
    except KeyboardInterrupt:
        print("Closing loop.")



def getData():
    END_OF_LIST = (1644, 649, 60, 53)

    # clear file
    with open("inboxErrors.csv", "w") as err:
        err.write("")

    cfg = None
    def first_time_setup():
        init = SimpleNamespace()
        init.rows = ['Material', 'Wbs', 'Plant', 'Qty']

        # if container tab is not selected
        try:
            x, y = pyautogui.locateCenterOnScreen(r"inboxImg\ContainerUnselected.PNG", grayscale=True, confidence=0.75)
            pyautogui.click(x, y)
        except TypeError:
            assert pyautogui.locateOnScreen(r"inboxImg\ContainerSelected.PNG", grayscale=True) is not None

        # expand horizontal divider up so that values all show when expanded
        x, y = pyautogui.locateCenterOnScreen(r"inboxImg\HorizontalDivider.PNG", grayscale=True)
        if y > 400:
            pyautogui.moveTo(x, y)
            pyautogui.dragTo(x, 400, button='left')

        # find input data expand button
        x, y, w, h = pyautogui.locateOnScreen(r"inboxImg\SigmanestInputData.PNG", grayscale=True, confidence=0.75)
        init.input_data_expand = dict(x=x + 5, y=y + 10)

        pyautogui.click(**init.input_data_expand)
        
        # get column start x-position
        x, y, w, h = pyautogui.locateOnScreen(r"inboxImg\ValuesHeader.PNG", grayscale=True, confidence=0.75)
        init.col_start = x + 5

        for i, r in enumerate(init.rows):
            x, y, w, h = pyautogui.locateOnScreen("inboxImg\\InputData{}.PNG".format(r), grayscale=True, confidence=0.75)
            init.rows[i] = (y, h)

        return init


    findOnScreen(r"inboxImg\InboxHeader.PNG")
    pyautogui.click(LINE_ITEM_CLICK)
    while whileLoopRun:
        pyautogui.click(findOnScreen(r"inboxImg\WorkflowLog.PNG"))

        # wait until log screen
        findOnScreen(r"inboxImg\WorkflowLogHeader.PNG")
        time.sleep(1)
        
        if not cfg:
            cfg = first_time_setup()
        else:
            pyautogui.click(**cfg.input_data_expand)
        

        with futures.ThreadPoolExecutor() as executor:
            threads = list()
            for row_y, row_h in cfg.rows:
                region = (cfg.col_start, row_y, 500, row_h)
                threads.append(executor.submit(captureRow, region))

            # wait for threads to complete
            futures.wait(threads)

            it = iter(threads)
            matl = cleanUpInput(next(it).result())
            wbs = cleanUpInput(next(it).result())
            plant = cleanUpInput(next(it).result())
            qty = cleanUpInput(next(it).result(), doNotRemove=["."])

        items = ",".join((matl, wbs, plant, qty))
        with open("inboxErrors.csv", "a") as err:
            err.write(items + "\n")
        print(items)

        pyautogui.press("f3")
        findOnScreen(r"inboxImg\InboxHeader.PNG")

        # check if last item (scrollbar at the bottom)
        if pyautogui.locateOnScreen(r"inboxImg\endOfList.PNG", grayscale=True):
            break

        pyautogui.press("down")
        time.sleep(0.25)


def thread_worker(index, x, img):
    _, y, _, h = findOnScreen(img, center=False)

    keep = []
    if index == "qty":
        keep = ["."]
    return index, cleanUpInput(captureRow((x, y, 500, h)), doNotRemove=keep)


def cleanUpInput(input, doNotRemove=[], cleanType=None):
    input = input.strip()

    REMOVE = ["/", "|", "_", ",", ".", "'", "�", "‘", "\\", "°", "("]
    REPLACE = [("$", "S"), ("§", "S"), ("S1", "S-1"), ("D1", "D-1")]
    for x in REMOVE:
        if x not in doNotRemove:
            input = input.replace(x, "")

    input = input.upper()
    for f, r in REPLACE:
        input = input.replace(f, r)

    if input in ('HSO1', '4501'):
        input = 'HS01'
    elif input in ('HSO2', '4502'):
        input = 'HS02'

    return input


def findAtLocation(picture, region):
    original = numpy.array(Image.open(picture))

    # start = time.time()
    while 1:
        current = numpy.array(pyautogui.screenshot(region=region))
        if numpy.max(numpy.abs(original - current)) == 0:
            break

        # if runs from more than 5 seconds,
        # overwrite region with find on screen
        # if time.time() - start > 5:
        #     region = pyautogui.locateOnScreen(picture)


def findOnScreen(picture, region=None, center=True):
    if center:
        locate = pyautogui.locateCenterOnScreen
    else:
        locate = pyautogui.locateOnScreen

    while True:
        try:
            res = locate(picture, grayscale=True)

            if res:
                return res
        except pyautogui.ImageNotFoundException:
            pass


def captureRow(region):
    f = "temp\\{}.png".format(uuid.uuid4())
    capture = pyautogui.screenshot(f, region=region)

    return pytesseract.image_to_string(capture, config="--oem 3 --psm 6")


def testLoc(img):
    loc = pyautogui.locateCenterOnScreen(img, grayscale=True)
    print(loc)

    if loc:
        pyautogui.moveTo(*loc)

        print("located")


if __name__ == '__main__':
    main()

    # time.sleep(2)
    # testLoc(r"inboxImg\ContainerUnselected.PNG")
    # testLoc(r"inboxImg\WorkflowLogHeader.PNG")

    # matl = captureRow(MATL_VALUE_REGION)
    # wbs = captureRow(WBS_VALUE_REGION)
    # qty = captureRow(QTY_VALUE_REGION)
    # print(matl, "::", wbs, "::", qty)
