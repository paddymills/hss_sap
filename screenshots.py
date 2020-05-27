import os
import re

from types import SimpleNamespace

class TransactonSet:

    def addOp(self, transaction, name, path):
        if not hasattr(self, transaction):
            setattr(self, transaction, OperationSet(name, path))
        else:
            setattr(getattr(self, transaction), name, path)

class OperationSet:

    def __init__(self, name, path):
        self.addOp(name, path)

    def addOp(self, name, path):
        setattr(self, name, path)

SCREENSHOT_RE = re.compile(r"([a-zA-Z0-9]+)_([a-zA-Z0-9]+).PNG")

def getScreenShotCollection():
    screenshots = TransactonSet()
    for f in os.scandir(os.path.join(os.path.dirname(__file__), "screenshots")):
        # transaction, name = f.name.replace(".PNG", "").split("_")
        transaction, name = SCREENSHOT_RE.match(f.name).groups()
        screenshots.addOp(transaction, name, f.path)

    return screenshots
