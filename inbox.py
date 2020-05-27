import pyautogui
import time
import os
import sys

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
    argFunctions = {
        "get": getData,
        "manual": manuallyEnterWBS,
        "fullauto": fullyAutomated,
        "test": testLoc,
    }

    try:
        if len(sys.argv) > 1:
            for arg in sys.argv[1:]:
                if arg in argFunctions.keys():
                    argFunctions[arg]()
                elif arg == "help":
                    print(
                        "available arguments:\n  ::" +
                        "\n  ::".join(argFunctions.keys())
                    )
        else:
            # fullyAutomated()
            # manuallyEnterWBS()
            getData()
    except KeyboardInterrupt:
        print("Closing loop.")


def manuallyEnterWBS():
    findAtLocation(r"inboxImg\InboxHeader.PNG", region=INBOX_HEADER_REGION)
    pyautogui.click(LINE_ITEM_CLICK)
    while whileLoopRun:
        # findAtLocation(r"inboxImg\LineText.PNG", region=LINE_ITEM_REGION)
        findAtLocation(r"inboxImg\WorkflowLog.PNG",
                       region=WORKFLOW_LOG_ICON_REGION)
        pyautogui.click(WORKFLOW_LOG_CLICK)

        findAtLocation(r"inboxImg\WorkflowLogHeader.PNG",
                       region=WORKFLOW_LOG_HEADER_REGION)
        pyautogui.click(EXPAND_INPUT_DATA_CLICK)

        matl = cleanUpInput(captureRow(MATL_VALUE_REGION), cleanType="job")
        wbs = cleanUpInput(captureRow(WBS_VALUE_REGION), cleanType="wbs")
        qty = cleanUpInput(captureRow(QTY_VALUE_REGION), doNotRemove=["."])

        print(matl, "::", wbs, "::", qty)
        pyautogui.press("f3")

        findAtLocation(r"inboxImg\InboxHeader.PNG", region=INBOX_HEADER_REGION)
        pyautogui.click(EXECUTE_CLICK)
        time.sleep(0.25)
        pyautogui.typewrite("D-" + matl[:7] + "-")
        time.sleep(2)


def fullyAutomated():
    homePath = os.path.dirname(os.path.realpath(__file__))
    with open(os.path.join(homePath, "parts.csv"), "r") as f:
        # line format: [part],[qty],[wbs],[shipment]
        parts = [x.replace("\n", "").split(",") for x in f.readlines()]
    with open(os.path.join(homePath, "inboxPartConv.csv"), "r") as f:
        conv = dict()
        for x in f.readlines():
            x1, x2 = x.replace("\n", "").split(",")
            conv[x1] = x2

    while whileLoopRun:
        findAtLocation(r"inboxImg\InboxHeader.PNG", region=INBOX_HEADER_REGION)
        pyautogui.click(LINE_ITEM_CLICK)

        findAtLocation(r"inboxImg\LineText.PNG", region=LINE_ITEM_REGION)
        findAtLocation(r"inboxImg\WorkflowLog.PNG",
                       region=WORKFLOW_LOG_ICON_REGION)
        pyautogui.click(WORKFLOW_LOG_CLICK)

        findAtLocation(r"inboxImg\WorkflowLogHeader.PNG",
                       region=WORKFLOW_LOG_HEADER_REGION)
        pyautogui.click(EXPAND_INPUT_DATA_CLICK)

        matl = cleanUpInput(captureRow(MATL_VALUE_REGION), cleanType="job")
        wbs = cleanUpInput(captureRow(WBS_VALUE_REGION), cleanType="wbs")
        qty = cleanUpInput(captureRow(QTY_VALUE_REGION), doNotRemove=["."])

        matl = conv.get(matl, matl)  # get matl if in conv, else matl

        print(matl, "::", wbs, "::", qty, end="")
        pyautogui.press("f3")

        findAtLocation(r"inboxImg\InboxHeader.PNG", region=INBOX_HEADER_REGION)
        pyautogui.click(EXECUTE_CLICK)

        qty = int(float(qty))
        shipment = wbs[-2:]
        potential = []
        for i, x in enumerate(parts):
            _part, _qty, _wbs, _shipment = x
            if matl == _part and _shipment == shipment:
                if int(_qty) > qty:
                    potential.append(x)
                else:
                    parts[i][1] = str(int(parts[i][1]) - qty)
                    update_wbs = _wbs
                    press = "enter"
                    print(" [updated]")
                    break
        else:
            print(" [not found]")
            update_wbs = FAILURE_STRING
            press = "f12"
            with open("failures.csv", "a") as f:
                f.write(",".join([matl, wbs, str(qty)]) + "\n")

        time.sleep(0.5)
        pyautogui.typewrite(update_wbs)
        pyautogui.press(press)
        time.sleep(1)


def getData():
    END_OF_LIST = (1644, 649, 60, 53)

    # clear file
    with open("inboxErrors.csv", "w") as err:
        err.write("")

    findAtLocation(r"inboxImg\InboxHeader.PNG", region=INBOX_HEADER_REGION)
    pyautogui.click(LINE_ITEM_CLICK)
    while whileLoopRun:
        findAtLocation(r"inboxImg\WorkflowLog.PNG",
                       region=WORKFLOW_LOG_ICON_REGION)
        pyautogui.click(WORKFLOW_LOG_CLICK)

        findAtLocation(r"inboxImg\WorkflowLogHeader.PNG",
                       region=WORKFLOW_LOG_HEADER_REGION)
        pyautogui.click(EXPAND_INPUT_DATA_CLICK)

        matl = cleanUpInput(captureRow(MATL_VALUE_REGION), cleanType="job")
        wbs = cleanUpInput(captureRow(WBS_VALUE_REGION), cleanType="wbs")
        qty = cleanUpInput(captureRow(QTY_VALUE_REGION), doNotRemove=["."])

        pyautogui.press("f3")
        findAtLocation(r"inboxImg\InboxHeader.PNG", region=INBOX_HEADER_REGION)

        with open("inboxErrors.csv", "a") as err:
            err.write(f"{matl},{wbs},{qty}\n")
        print(f"{matl},{wbs},{qty}")

        original = numpy.array(Image.open(r"inboxImg\endOfList.PNG"))
        current = numpy.array(pyautogui.screenshot(region=END_OF_LIST))
        if numpy.max(numpy.abs(original - current)) == 0:
            break

        pyautogui.press("down")
        time.sleep(0.25)


def cleanUpInput(input, doNotRemove=[], cleanType=None):
    REMOVE = ["/", "|", "_", ",", ".", "'", "�", "‘"]
    for x in REMOVE:
        if x not in doNotRemove:
            input = input.replace(x, "")

    if cleanType == "job":
        return input[:8] + input[input.find("-"):]
    elif cleanType == "wbs":
        return input.replace("\\", "").replace("s1", "S-1").replace("S1", "S-1")

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


def captureRow(region):
    capture = pyautogui.screenshot(region=region)
    return pytesseract.image_to_string(capture)


def testLoc():
    picture = r"inboxImg\LineText.PNG"
    region = LINE_ITEM_REGION
    foundAt = (0, 0, 0, 0)

    # original = numpy.array(Image.open(picture))
    # while 1:
    #     current = numpy.array(pyautogui.screenshot(region=region))
    #     if numpy.max(numpy.abs(original - current)) == 0:
    #         break

    #     _region = pyautogui.locateOnScreen(picture)
    #     if foundAt != _region:
    #         foundAt = _region
    #         print(f"x:{region[0] - foundAt[0]} y:{region[1] - foundAt[1]}")

    pyautogui.moveTo(EXPAND_INPUT_DATA_CLICK)

    print("located")


if __name__ == '__main__':
    main()

    # matl = captureRow(MATL_VALUE_REGION)
    # wbs = captureRow(WBS_VALUE_REGION)
    # qty = captureRow(QTY_VALUE_REGION)
    # print(matl, "::", wbs, "::", qty)
