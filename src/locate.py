import pyautogui
import time


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
    pyautogui.screenshot(name, region=(x1, y1, x2-x1, y2-y1))
    print("width:", x2-x1, "  height:", y2-y1)


def takePictureCoordAndSize(x1, y1, wid, height, name="image.PNG"):
    pyautogui.screenshot(name, region=(x1, y1, wid, height))


if __name__ == '__main__':
    # findCoord()

    time.sleep(5)
    takePictureTwoCoord(20, 150, 110, 195)
    # takePictureCoordAndSize(25, 82, 475, 18)
    print('done')
