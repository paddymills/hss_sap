import pyautogui
import time

import sys

def findCoord():
    DELAY = 0.25
    last_x, last_y, last_rgb = None, None, None

    while 1:
        x, y = pyautogui.position()
        rgb = pyautogui.pixel(x, y)

        if x == last_x and y == last_y and rgb == last_rgb:
            continue
        _x, _y = (str(x).rjust(4) for x in pyautogui.position())
        print("x:", _x, "  y:", _y, rgb)
        last_x = x
        last_y = y
        last_rgb = rgb

        # time.sleep(DELAY)


def takePictureTwoCoord(x1, y1, x2, y2, name="image.PNG"):
    takePictureCoordAndSize(x1, y1, x2-x1, y2-y1, name=name)


def takePictureCoordAndSize(x1, y1, wid, height, name="image.PNG"):
    pyautogui.screenshot(name, region=(x1, y1, wid, height))
    print("Screenshot taken:", name)
    print("location tuple:", (x1, y1, wid, height))


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == "take":
        for x in range(5, 0 ,-1):
            print("\rPicture in :", x, end="")
            time.sleep(1)
        print("")
        
        loc = (68, 74, 424, 93)
        name = "CO02_OperationOverviewHeader"

        # takePictureCoordAndSize(*loc, name="screenshots\\{}.PNG".format(name))
        takePictureTwoCoord(*loc, name="screenshots\\{}.PNG".format(name))
        
    else:
        findCoord()
