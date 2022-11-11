from system_interfaces import *  # also imports most libraries
from ctypes import windll
from time import sleep
import pytesseract
import re
from win32api import keybd_event
import os

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def updateScreenshot(hwnd):
    screenshot = get_screenshot(hwnd)
    screenshot = np.array(screenshot)
    return cv.cvtColor(screenshot, cv.COLOR_RGB2BGR)


def click_at_center(rectangle_):
    click(get_center_of_rectangle(rectangle_)[0], get_center_of_rectangle(rectangle_)[1] + 20, rectangle_)


def find_all(txt, substr: str):
    return [m.start() for m in re.finditer(substr, txt)]


def KeyUp(Key):
    keybd_event(Key, 0, 2, 0)


def KeyDown(Key):
    keybd_event(Key, 0, 1, 0)


def press(*args):
    """
    one press, one release.
    accepts as many arguments as you want. e.g. press('left_arrow', 'a','b').
    """
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, 0, 0)
        time.sleep(.05)
        win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)


def main():
    windll.user32.SetProcessDPIAware()
    hwnd = get_hwnd("brave")
    # print(hwnd)
    win32gui.SetForegroundWindow(hwnd)
    sleep(0.1)
    rectangle = win32gui.GetWindowRect(hwnd)
    # click_at_center(rectangle)
    # make sure we are at the top
    for i in range(10):
        scroll(True, 1000)

    sleep(5)
    text = ""
    while not text.find("Send me") > 0:
        screenshot = get_screenshot(hwnd)
        temp = pytesseract.image_to_string(screenshot)
        text += temp
        scroll(False, 500)

    try:
        os.remove("out.txt")
    except:
        pass
    with open("out.txt", "w") as file:
        all = find_all(text, "below?")
        for index in all:
            temp = text[index + 6:]
            temp = temp[:temp.find("What")]
            # print(temp)
            # print("-------------")
            file.write(temp)
            file.write("---------------------------------")

main()
