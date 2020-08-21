import pyautogui
import xlwings
import time
import re
import os
import sys
from datetime import date

from multiprocessing import Pool, Process, Queue, current_process
import queue
import tqdm

from PIL import Image, ImageGrab
import pytesseract
import numpy

import screenshots

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"
img = screenshots.getScreenShotCollection()

ITEM_REGEX = re.compile("^[0-9]{4}$")
TIME_REGEX = re.compile("^[0-9]{2}:[0-9]{2}:[0-9]{2}$")

ACTIVE_LINE_RGB = (255, 255, 255)
DISABLED_LINE_RGB = (223, 235, 245)

SAP_TABLE_LINE_HEIGHT = 21
CO02_TABLE_LEFT = 35
CO02_TABLE_TOP = 282
CO02_TABLE_BOTTOM = 891

CO02_ITEM_WIDTH_OPERATIONS = 563
CO02_ITEM_WIDTH_COMPONENTS = 150

CO02_YES_BUTTON = (499, 288, 33, 31)
CO02_ORDER_LABEL_BUTTON = (18, 193, 115, 32)

EMPTY_CELL_REPLACEMENTS = ["", "="]


def main():
    argFunctions = {
        "remove": helpRemoveLines,
        "manual": manuallyAddOperationsAndComponents,
        "add": manuallyAddOperationsAndComponents,
        "check": checkWinShuttleOrManual,
        "unconfirm": helpUnConfirm0444,
        "unconfirm_part": helpUnConfirmPart,
        "delete": setDeletionFlag,
        "mrp": runMRP,
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
            # co02AddMaterialLineCsv()
            # checkWinShuttleOrManual()
            # manuallyAddOperationsAndComponents(plant="HS02")  # plant is hardcoded
            # helpRemoveLines()
            helpUnConfirm0444(query0444=False)
    except KeyboardInterrupt:
        print("Closing loop.")


# arg: "check"
def checkWinShuttleOrManual():
    CAPTURE_REGIONS = []
    def region(x): return (CO02_TABLE_LEFT, x,
                           CO02_ITEM_WIDTH_OPERATIONS, SAP_TABLE_LINE_HEIGHT)
    for i in range(CO02_TABLE_TOP, CO02_TABLE_BOTTOM, SAP_TABLE_LINE_HEIGHT):
        CAPTURE_REGIONS.append(region(i))

    INITIAL_SCREEN_HEADER = (25, 80, 350, 25)
    INITIAL_SCREEN_ORDER = (168, 204)
    OPERATIONS_HEADER = (75, 82, 405, 20)
    BACK_BUTTON = (282, 45, 15, 15)

    orders = read_sort_min_file("orders.txt")

    iterations = len(orders)
    with Pool() as p:
        for i, order in enumerate(orders, start=1):
            loopFunc(findAtLocation, img.CO02.InitialScreenHeader,
                     INITIAL_SCREEN_HEADER)
            pyautogui.click(*INITIAL_SCREEN_ORDER)
            pyautogui.typewrite(order)
            pyautogui.press("f5")

            loopFunc(findAtLocation, img.CO02.OperationOverviewHeader,
                     OPERATIONS_HEADER)
            pyautogui.click(x=120, y=160)
            processed = list(tqdm.tqdm(p.imap(checkOperationsLine, CAPTURE_REGIONS),
                                       desc=f"{i}/{iterations} :: {order}",
                                       total=len(CAPTURE_REGIONS)))
            if any(processed):
                print("\t^-> operation added")
            pyautogui.click(x=290, y=50)


def checkOperationsLine(region):
    y, text = captureRow(region)
    if "MATLCONS" in text.upper():
        return True
    return None


# arg: "remove"
def helpRemoveLines():
    def region(x): return (CO02_TABLE_LEFT, x,
                           CO02_ITEM_WIDTH_OPERATIONS, SAP_TABLE_LINE_HEIGHT)

    def getCaptureRegions():
        q = Queue()
        for i in range(CO02_TABLE_TOP, CO02_TABLE_BOTTOM, SAP_TABLE_LINE_HEIGHT):
            q.put(region(i))
        return q

    INITIAL_SCREEN_HEADER = (25, 80, 350, 25)
    INITIAL_SCREEN_ORDER = (168, 204)
    OPERATIONS_HEADER = (75, 82, 405, 20)
    BACK_BUTTON = (282, 45, 15, 15)
    CONFIRM_YES = (442, 288, 153, 26)
    CALC_COSTS = (234, 447, 71, 16)

    AVAILABLE_PROCESSES = os.cpu_count()

    orders = read_sort_min_file("orders.txt")

    total = len(orders)
    for i, order in enumerate(orders, start=1):
        loopFunc(findAtLocation, img.CO02.InitialScreenHeader,
                 INITIAL_SCREEN_HEADER)
        pyautogui.click(*INITIAL_SCREEN_ORDER)
        pyautogui.typewrite(order)
        pyautogui.press("f5")

        loopFunc(findAtLocation, img.CO02.OperationOverviewHeader,
                 OPERATIONS_HEADER)
        # move cursor out of A1 box (otherwise can cause Tesseract misread)
        pyautogui.click(x=120, y=160)

        taskQueue = getCaptureRegions()
        doneQueue = Queue()
        shutdownQueue = Queue()
        lineSelectClicks = 0
        with tqdm.tqdm(desc=f"{i}/{total} :: {order}", total=taskQueue.qsize()) as progress:
            processes = dict()
            for p in range(AVAILABLE_PROCESSES):
                processes[p] = Process(
                    target=captureWorker,
                    args=(p, taskQueue, doneQueue, shutdownQueue)
                )
                processes[p].start()

            while any([x.is_alive() for x in processes.values()]):
                # shutdown any processes that hit an empty queue
                while not shutdownQueue.empty():
                    id = shutdownQueue.get_nowait()
                    processes[id].terminate()

                try:  # needed to keep doneQueue.get() from blocking
                    process, y, text = doneQueue.get_nowait()
                except queue.Empty:
                    continue

                progress.update(1)
                items = text.replace("\n", " ").split(" ")
                if len(items) == 3:  # empty row: clear queue
                    try:
                        while True:
                            taskQueue.get_nowait()
                    except queue.Empty:
                        pass
                # check if line is active
                elif pyautogui.pixelMatchesColor(CO02_TABLE_LEFT+15, y+10, ACTIVE_LINE_RGB):
                    if "MATLCONS" in items:
                        pyautogui.click(x=CO02_TABLE_LEFT-10, y=y)
                        lineSelectClicks += 1

        # delete selected lines
        if lineSelectClicks > 0:
            pyautogui.click(x=124, y=995)  # delete button
            loopFunc(findAtLocation, img.CO02.ConfirmYes, CONFIRM_YES)
            pyautogui.click(pyautogui.center(CONFIRM_YES))
            time.sleep(0.5)
            pyautogui.hotkey("ctrl", "s")
        else:
            pyautogui.hotkey("shift", "f3")

        while 1:
            if findAtLocation(img.CO02.InitialScreenHeader, INITIAL_SCREEN_HEADER):
                break
            elif findAtLocation(img.CO02.CalculatingCosts, CALC_COSTS):
                pyautogui.click(pyautogui.center(CALC_COSTS))
                break


# arg: "manual"
# arg: "add"
def manuallyAddOperationsAndComponents(plant=None):
    if plant is None:
        plant = input("Plant: ")

    if plant in ('1', '2', '4'):
        plant = 'HS01'
    elif plant == '3':
        plant = 'HS02'

    CAPTURE_REGIONS = list()
    def region(x): return (CO02_TABLE_LEFT, x,
                           CO02_ITEM_WIDTH_OPERATIONS, SAP_TABLE_LINE_HEIGHT)
    for i in range(CO02_TABLE_TOP, CO02_TABLE_BOTTOM, SAP_TABLE_LINE_HEIGHT):
        CAPTURE_REGIONS.append(region(i))

    INITIAL_SCREEN_HEADER = (25, 80, 350, 25)
    INITIAL_SCREEN_ORDER = (168, 204)
    OPERATION_HEADER = (75, 82, 405, 20)
    COMPONENT_HEADER = (75, 83, 420, 22)
    BACK_BUTTON = (282, 45, 15, 15)
    BLANK_CELL = (71, 306, 120, 8)
    ERROR_CALC = (235, 448, 71, 14)

    WINSHUTTLE_SUCCESS_REGEX = re.compile("^Order number [0-9]{10} saved$")

    data = dict()
    deleted0444 = list()
    order = None
    wb = xlwings.books.active

    starting_row = 2
    if len(sys.argv) > 2:
        try:
            starting_row = int(sys.argv[2])
        except:
            pass

    rng = "A{}:D{}".format(starting_row, starting_row)
    for x in wb.sheets["Sheet1"].range(rng).expand('down').value:
        if x[3] and WINSHUTTLE_SUCCESS_REGEX.match(x[3]):
            continue

        if x[0] != order:
            order = x[0]
            data[order] = []
        if x[3] == "Selection of deleted operations not allowed":
            deleted0444.append(order)
        data[order].append([x[1], str(int(x[2]))])

    progress = tqdm.tqdm(total=len(data.keys()))
    for order, parts in data.items():
        if order in deleted0444:
            lineNumber = "0454"
        else:
            lineNumber = "0444"

        progress.update(1)
        loopFunc(findAtLocation, img.CO02.InitialScreenHeader,
                 INITIAL_SCREEN_HEADER)
        pyautogui.click(*INITIAL_SCREEN_ORDER)
        pyautogui.typewrite(order)
        pyautogui.press("f5")

        loopFunc(findAtLocation, img.CO02.OperationOverviewHeader,
                 OPERATION_HEADER)
        pyautogui.press("pagedown")
        while not pyautogui.pixelMatchesColor(83, 313, ACTIVE_LINE_RGB):
            pass
        pyautogui.click(x=38, y=314)

        # add Operation
        region = (CO02_TABLE_LEFT,
                  CO02_TABLE_TOP,
                  CO02_ITEM_WIDTH_OPERATIONS,
                  SAP_TABLE_LINE_HEIGHT)
        if not checkOperationsLine(region):
            for x in (lineNumber, None, "MATLCONS", None, "ZP01", "2034"):
                if x is not None:
                    pyautogui.typewrite(x)
                pyautogui.press("tab")
        pyautogui.press("f6")
        while not findAtLocation(img.CO02.ComponentOverviewHeader, COMPONENT_HEADER):
            pyautogui.press("enter")

        time.sleep(0.5)
        pyautogui.press("pagedown")
        loopFunc(findAtLocation, img.SAP.BlankCell, BLANK_CELL)

        # add components
        for i, partandqty in enumerate(parts):
            pyautogui.click(x=71, y=i*SAP_TABLE_LINE_HEIGHT+314)
            for x in (*partandqty, "EA", "L", lineNumber, "0", plant, "PROD"):
                pyautogui.typewrite(x)
                pyautogui.press("tab")
        pyautogui.hotkey("ctrl", "s")
        while 1:
            if findAtLocation(img.CO02.InitialScreenHeader, INITIAL_SCREEN_HEADER):
                break
            elif findAtLocation(img.CO02.ErrorCalculatingCostsYes, ERROR_CALC):
                pyautogui.click(pyautogui.center(ERROR_CALC))
                break


# arg: "unconfirm"
def helpUnConfirm0444(query0444=True):
    INITIAL_SCREEN_HEADER = (27, 83, 455, 18)
    INITIAL_SCREEN_ORDER = (140, 271)
    DETAILS_SCREEN_HEADER = (25, 82, 435, 18)
    CONF_TEXT_SCREEN_HEADER = (25, 82, 435, 18)
    CONF_SELECT_DIALOG_TITLE = (75, 190, 352, 25)
    UNCONFIRM_DATE = (146, 664)

    TODAY = date.today().strftime('%m/%d/%Y')

    orders = read_sort_min_file("orders.txt")

    def cancel_conf(order):
        loopFunc(findAtLocation, img.CO13.InitialScreenHeader,
                 INITIAL_SCREEN_HEADER)
        pyautogui.click(*INITIAL_SCREEN_ORDER)
        pyautogui.typewrite(order)
        if query0444:
            pyautogui.press("tab")
            pyautogui.typewrite("0")
            pyautogui.press("tab")
            pyautogui.typewrite("0444")
        pyautogui.press("enter")
        progress.write(f"Unconfirming {order}")
        time.sleep(.25)
        pyautogui.press("enter")

        while 1:
            if findAtLocation(img.CO13.Details, DETAILS_SCREEN_HEADER):
                rerun_order = False
                break
            elif findAtLocation(img.CO13.ConfirmationSelection, CONF_SELECT_DIALOG_TITLE):
                rerun_order = True
                break
        loopFunc(findAtLocation, img.CO13.Details,
                 DETAILS_SCREEN_HEADER)
        pyautogui.click(*UNCONFIRM_DATE)
        pyautogui.typewrite(TODAY)
        pyautogui.hotkey('ctrl', 's')
        loopFunc(findAtLocation, img.CO13.ConfirmationText,
                 CONF_TEXT_SCREEN_HEADER)
        pyautogui.press('escape')

        if rerun_order:
            cancel_conf(order)

    with tqdm.tqdm(orders) as progress:
        for order in progress:
            cancel_conf(order)


# arg: "unconfirm_part"
def helpUnConfirmPart():
    INITIAL_SCREEN_HEADER = (27, 83, 455, 18)
    INITIAL_SCREEN_ORDER = (140, 271)
    CONF_ORDER_TITLE = (64, 212, 225, 28)
    DETAILS_SCREEN_HEADER = (25, 82, 475, 18)
    CONF_TEXT_SCREEN_HEADER = (25, 82, 435, 18)

    TODAY = date.today().strftime('%m/%d/%Y')

    orders = read_sort_min_file("orders.txt")

    def cancel_conf(order):
        loopFunc(findAtLocation, img.CO13.InitialScreenHeader,
                 INITIAL_SCREEN_HEADER)
        pyautogui.click(*INITIAL_SCREEN_ORDER)
        pyautogui.typewrite(order)
        pyautogui.press("enter")
        progress.write(f"Unconfirming {order}")
        loopFunc(findAtLocation, img.CO13.ConfirmOrder,
                 CONF_ORDER_TITLE)
        pyautogui.press("enter")

        loopFunc(findAtLocation, img.CO13.DetailsActualData,
                 DETAILS_SCREEN_HEADER)
        pyautogui.typewrite(TODAY)
        pyautogui.hotkey('ctrl', 's')
        loopFunc(findAtLocation, img.CO13.ConfirmationText,
                 CONF_TEXT_SCREEN_HEADER)
        pyautogui.press('escape')

    with tqdm.tqdm(orders) as progress:
        for order in progress:
            cancel_conf(order)


# arg: "delete"
def setDeletionFlag():
    INITIAL_SCREEN_HEADER = (25, 80, 350, 25)
    INITIAL_SCREEN_ORDER = (168, 204)

    orders = read_sort_min_file("orders.txt")

    with tqdm.tqdm(orders) as progress:
        for order in progress:
            progress.write("Setting deletion flag for {}".format(order))
            loopFunc(findAtLocation, img.CO02.InitialScreenHeader,
                     INITIAL_SCREEN_HEADER)
            pyautogui.click(*INITIAL_SCREEN_ORDER)
            pyautogui.typewrite(order)
            pyautogui.press("enter")

            time.sleep(.25)

            # delete selected lines
            pyautogui.keyDown("alt")
            pyautogui.press("n")
            pyautogui.press("l")
            pyautogui.press("s")
            pyautogui.keyUp("alt")
            time.sleep(.5)
            pyautogui.hotkey("ctrl", "s")


# arg: "mrp"
def runMRP():
    PROJECT = (196, 160)
    WBS = (196, 182)
    MRP_MAIN_PAGE = (20, 150, 90, 45)
    SUCCESS_COLORS = (300, 145, 100, 95)

    PROJECT_RE = re.compile("^[a-zA-Z]-[0-9]{7}$")
    WBS_RE = re.compile("^[a-zA-Z]-[0-9]{7}-[0-9]{5}$")

    data = read_sort_min_file("orders.txt")

    with tqdm.tqdm(data) as progress:
        for project_or_wbs in progress:
            loopFunc(findAtLocation, img.MB51.MainPageKey, MRP_MAIN_PAGE)
            # enter Project or WBS
            if PROJECT_RE.match(project_or_wbs):
                pyautogui.click(*PROJECT)
            elif WBS_RE.match(project_or_wbs):
                pyautogui.click(*WBS)
            else:
                continue

            pyautogui.typewrite(project_or_wbs)
            pyautogui.press("enter")
            time.sleep(.5)
            pyautogui.press("enter")

            loopFunc(findAtLocation, img.MB51.SuccessColors, SUCCESS_COLORS)
            pyautogui.press("f3")


def co02AddMaterialLineCsv():
    CAPTURE_REGIONS = []
    def region(x): return (CO02_TABLE_LEFT, x,
                           CO02_ITEM_WIDTH_COMPONENTS, SAP_TABLE_LINE_HEIGHT)
    for i in range(CO02_TABLE_TOP, CO02_TABLE_BOTTOM, SAP_TABLE_LINE_HEIGHT):
        CAPTURE_REGIONS.append(region(i))
    WINSHUTTLE_SUCCESS_REGEX = re.compile("^Order number [0-9]{10} saved$")

    with open("CO02 Add Material Line.csv", "r") as f:
        winshuttle_data = f.readlines()

    for row in winshuttle_data:
        order, result = row.strip().split(",")
        if WINSHUTTLE_SUCCESS_REGEX.match(result):
            removeBomItems(order, CAPTURE_REGIONS)


def removeBomItems(order, captureRegions):
    waitUntilPresent(img.CO02.OrderLabel)
    pyautogui.click(x=168, y=203)
    pyautogui.typewrite(order)
    locateAndClick(img.CO02.ComponentOverviewButton)
    waitUntilPresent(img.CO02.ComponentOverviewDisabledButton)

    remove = list()
    continueLoop = True
    while continueLoop:
        remove = []
        with Pool() as p:
            processed = list(tqdm.tqdm(
                p.imap(captureRow, captureRegions), desc=order, total=len(captureRegions)))
        for x in processed:
            items = x[1].replace("\n", " ").split(" ")

            if len(items) == 1:
                if ITEM_REGEX.match(items[0]) or len(items[0]) == 4:
                    continueLoop = False
                    break
                else:
                    if pyautogui.pixelMatchesColor(CO02_TABLE_LEFT+15, x[0]+10, ACTIVE_LINE_RGB):
                        remove.append(x[0] + 10)

        for y in remove:
            pyautogui.click(x=CO02_TABLE_LEFT-10, y=y)
        if remove:
            locateAndClick(img.SAP.DeleteButton)
            while 1:
                orderLabel = findAtLocation(
                    img.CO02.OrderLabel, CO02_ORDER_LABEL_BUTTON)
                promptYes = findAtLocation(img.CO02.PromptYes, CO02_YES_BUTTON)
                if orderLabel:
                    print("delete error.")
                    return
                elif promptYes:
                    pyautogui.click(promptYes)
                    break
            waitUntilPresent(img.SAP.SaveButton)

        if continueLoop:
            waitUntilScreenUpdated(locateAndClick(img.SAP.PageDown))
            pyautogui.click(x=503, y=290)  # select cursor out of data area

    locateAndClick(img.SAP.SaveButton)
    waitUntilPresent(img.CO02.ComponentOverviewButton)


def loopFunc(func, *args, **kwargs):
    while 1:
        returned = func(*args, **kwargs)
        if returned:
            break
        time.sleep(0.1)
    return returned


def findAtLocation(picture, region, original=None):
    if not original:
        original = numpy.array(Image.open(picture))

    current = numpy.array(pyautogui.screenshot(region=region))
    if numpy.max(numpy.abs(original - current)) == 0:
        return pyautogui.center(region)

    return None


def waitUntilScreenUpdated(initialAction, original=None):
    if not original:
        original = numpy.array(pyautogui.screenshot())
    while 1:
        current = numpy.array(pyautogui.screenshot())
        if numpy.max(numpy.abs(original - current)) != 0:
            return


def waitUntilPresent(screenshot):
    while 1:
        test = pyautogui.locateOnScreen(screenshot)
        if test is not None:
            return pyautogui.center(test)


def locateAndClick(screenshot):
    btn = waitUntilPresent(screenshot)
    pyautogui.click(x=btn.x, y=btn.y)


def iterateTable(**kwargs):
    tableLeft = kwargs["tableLeft"]
    tableTop = kwargs["tableTop"]
    tableWidth = kwargs["tableWidth"]
    tableBottom = kwargs["tableHeight"]
    rowHeight = kwargs["rowHeight"]

    def region(x): return (tableLeft, x, tableWidth, rowHeight)
    for i in range(tableTop, tableBottom, rowHeight):
        capture = pyautogui.screenshot(region=region(i))
        yield (i, pytesseract.image_to_string(capture))


def captureRow(region):
    capture = pyautogui.screenshot(region=region)
    return (region[1], pytesseract.image_to_string(capture))


def captureWorker(workerID, inputQueue, outputQueue, terminateQueue):
    try:
        for x in iter(inputQueue.get_nowait, "STOP"):
            outputQueue.put((workerID, *captureRow(x)))
    except queue.Empty:
        pass

    terminateQueue.put(workerID)


def read_sort_min_file(filename):
    homePath = os.path.dirname(os.path.realpath(__file__))

    # read file
    with open(os.path.join(homePath, filename), "r") as f:
        items = f.read().split("\n")

    # remove duplicates and sort
    orderd = sorted(set(items))

    # write sorted, minified list back to file
    with open(os.path.join(homePath, filename), "w") as f:
        f.writelines(orderd)

    return orderd


if __name__ == '__main__':
    main()
