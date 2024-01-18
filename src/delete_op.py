
import pyautogui
import pynput
import pytesseract
import os
import time
import tqdm

class MouseInputState:

    NeedsTrained = 0

    GetMainHeader1 = 1
    GetMainHeader2 = 2
    GetOrderInput = 3
    GetDeleteButton = 4

    GetErrorHeader1 = 5
    GetErrorHeader2 = 6

    AwaitingOpSelection = 50

    FullyTrained = 99

class ScreenImage:

    def __init__(self) -> None:
        self.start_coord = None
        self.size = None
        self.img = None

    def take_start(self, x, y):
        self.start_coord = (x, y)
    
    def take_end(self, x, y):
        self.size = (x - self.start_coord[0], y - self.start_coord[1])
        self.img = pyautogui.screenshot(region=(*self.start_coord, *self.size))

    def is_on_screen(self):
        try:
            coords = pyautogui.locateOnScreen(self.img)
            return coords == (*self.start_coord, *self.size)
        except pyautogui.ImageNotFoundException:
            return False
        
    def wait_until_visible(self):
        while not self.is_on_screen():
            time.sleep(0.25)

class OpDeleter:
    def __init__(self):
        self.input_state = MouseInputState.NeedsTrained

        self.order_loc = None
        self.main_header = ScreenImage()
        # self.op_overview_header = ScreenImage()
        self.error_header = ScreenImage()

        self.delete_button_loc = None
        self.error_yes_button_loc = None
        
        self.last_read_coord = None
        self.orders = read_sort_min_file("orders.txt")

        self.progress = None

    def run(self):
        print("Press Ctrl+C to quit")
        listener = pynput.mouse.Listener(on_click=self.handle_click)
        listener.start()
        try:
            listener.wait()
            self.train(self.orders.pop(0))
            
            # start processing
            self.progress = tqdm.tqdm(self.orders, desc="processing...")
            for order in self.progress:
                self.process(order)

        finally:
            listener.stop()


    def train(self, order):
        print("Open CO02 and select the start coordinate for the header")
        self.await_click()
        self.main_header.take_start(*self.last_read_coord)
        self.last_read_coord = None

        print("select the end coordinate for the header")
        self.await_click()
        self.main_header.take_end(*self.last_read_coord)
        self.last_read_coord = None

        print("select the start of the order box")
        self.await_click()
        self.order_loc = self.last_read_coord
        self.last_read_coord = None
        pyautogui.typewrite(order)
        pyautogui.press('f5')

        print("select the op to delete")
        self.await_click()
        self.last_read_coord = None

        print("select the delete op button")
        self.await_click()
        self.delete_button_loc = self.last_read_coord
        self.last_read_coord = None

        # print("select the start coordinate for the error header")
        # self.await_click()
        # self.error_header.take_start(*self.last_read_coord)
        # self.last_read_coord = None

        # print("select the end coordinate for the error header")
        # self.await_click()
        # self.error_header.take_end(*self.last_read_coord)
        # self.last_read_coord = None

        # print("select the coordinate for the yes button ")
        # self.await_click()
        # self.error_yes_button_loc = self.last_read_coord
        # self.last_read_coord = None
        # pyautogui.hotkey('ctrl', 's')
    

    def process(self, order):
        # self.main_header.wait_until_visible()
        time.sleep(5)

        pyautogui.click(*self.order_loc)
        pyautogui.typewrite(order)
        pyautogui.press('f5')
        
        self.progress.display("Select op to be deleted")
        self.await_click()
        self.last_read_coord = None

        # pyautogui.click(*self.delete_button_loc)
        time.sleep(5)

        # self.wait_until_visible(self.main_header, self.error_header)
        # if self.error_header.is_on_screen:
        #     pyautogui.click(*self.error_yes_button_loc)
        #     pyautogui.hotkey("ctrl", "s")

    def await_click(self):
        while self.last_read_coord is None:
            time.sleep(0.1)

    def wait_until_visible(self, *imgs):
        while not any([img.is_on_screen() for img in imgs]):
            time.sleep(0.25)


    def handle_click(self, x, y, button, pressed):
        # black list ops that are not left-click
        if not (button == pynput.mouse.Button.left and pressed):
            return

        self.last_read_coord = (x, y)


def read_sort_min_file(filename):
    homePath = os.path.dirname(os.path.realpath(__file__))

    # read file
    with open(os.path.join(homePath, filename), "r") as f:
        items = f.read().split("\n")

    # remove duplicates and sort
    orderd = sorted(set(items))
    if '' in orderd:
        orderd.remove('')

    # write sorted, minified list back to file
    with open(os.path.join(homePath, filename), "w") as f:
        f.write("\n".join(orderd))

    return orderd


if __name__ == "__main__":
    OpDeleter().run()
