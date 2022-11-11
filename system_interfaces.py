import pytesseract
from PIL import Image, ImageGrab
import cv2 as cv
import numpy as np
import time
import win32api
import win32con
import win32gui


def click(x: int, y: int, window_rectangle: list):
    """
    Clicks at the specified pixel coordinates relative to the window.
    :param x: x pixel
    :param y: y pixel
    :param window_rectangle: the list with the rectangle for the window
    :return:
    """
    win32api.SetCursorPos((x + window_rectangle[0], y + window_rectangle[1]))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def click_and_hold(x: int, y: int, hold_time: float, window_rectangle: list):
    """
    Clicks at the specified pixel coordinates relative to the window.
    :param x: x pixel
    :param y: y pixel
    :param hold_time: how long it holds the click for
    :param window_rectangle: the list with the rectangle for the window
    :return:
    """
    win32api.SetCursorPos((x + window_rectangle[0], y + window_rectangle[1]))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(hold_time)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def get_hwnd(window_title: str):
    """
    Gets the window handle of the specified application.
    :param window_title: the title of the application
    :return: the window handle
    """

    def enum_cb(hwnd, result):
        win_list.append((hwnd, win32gui.GetWindowText(hwnd)))

    top_list, win_list = [], []
    win32gui.EnumWindows(enum_cb, top_list)
    hwnd_list = [(hwnd, title) for hwnd, title in win_list if title.lower().find(window_title) > 0]

    if len(hwnd_list) != 0:
        return hwnd_list[0][0]
    return None


def get_screenshot(hwnd: int) -> Image:
    """
    Takes a screenshot of the specified application.
    :param hwnd: the window handle of the application.
    :return: the screenshot of the specified application
    """
    bbox = win32gui.GetWindowRect(hwnd)
    return ImageGrab.grab(bbox)


def find_image_rectangle(image_tuple: tuple, screenshot: Image) -> list:
    """
    Attempts to find the rectangle of an image in a screenshot, this is good for when the image only appears in one
    place on the screen at a time.
    :param image_tuple: a tuple for the image trying to be found containing (image, confidence)
    :param screenshot: the screenshot that the image might be in
    :return: a list with the rectangle for the image in [x, y, w, h] or [] if the image is not found
    """
    image, confidence = image_tuple
    threshold = 1 - confidence
    result = cv.matchTemplate(screenshot, image, cv.TM_SQDIFF_NORMED)
    location = np.where(result <= threshold)
    location = list(zip(*location[::-1]))
    rectangle = []
    if location:
        # Gets the first time the image is found
        location = location[0]
        rectangle = [location[0], location[1], image.shape[1], image.shape[0]]
    return rectangle


def scroll(up: bool, amount: int) -> None:
    if up:
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, abs(amount), 0)
    else:
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -abs(amount), 0)


def get_center_of_rectangle(rectangle: list) -> tuple:
    """
    Finds the center of a rectangle
    :param rectangle: the rectangle list [x, y, w, h]
    :return: a tuple containing the center in (x, y)
    """
    x = int(rectangle[0] + rectangle[2] / 2)
    y = int(rectangle[1] + rectangle[3] / 2)
    return x, y


def detect_if_color_present(color: list, cropped_screenshot: Image) -> bool:
    """
    Returns True if the specified color is in the screenshot
    :param color: the [r, g, b] list representing the color
    :param cropped_screenshot: the screenshot of the location where the color might be
    :return: a boolean for if the color is in the image
    """
    for y in range(len(cropped_screenshot)):
        for x in range(len(cropped_screenshot[y])):
            pixel_color = cropped_screenshot[y][x]
            if (abs(color[0] - pixel_color[0]) < 10 and
                    abs(color[1] - pixel_color[1]) < 10 and
                    abs(color[2] - pixel_color[2]) < 10):
                return True
    return False


def read_text(cropped_screenshot: Image) -> str:
    """
    Attempts to read the text in the cropped screenshot
    :param cropped_screenshot: the screenshot of the location where the text might be
    :return: the string of the text, "" if no text is found
    """
    pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Ryan\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\LocalCache\local-packages\Python39\Scripts\pytesseract.exe"
    processed_image = process_image_for_reading(cropped_screenshot)
    text = pytesseract.image_to_string(processed_image, lang="eng",
                                       config="-c tessedit_char_whitelist=0123456789 --psm 6")
    return text


def process_image_for_reading(cropped_screenshot: Image) -> Image:
    """
    Processes an image so that the text on it can be read easily
    :param cropped_screenshot: the image with text on it that needs to be processed
    :return: the processed image
    """
    # cv.imshow("B", cropped_screenshot)
    hsv = cv.cvtColor(cropped_screenshot, cv.COLOR_BGR2HSV)
    # define range of text color in HSV
    lower_value = np.array([0, 0, 100])
    upper_value = np.array([179, 120, 255])
    # filters the HSV image to get only the text color, returns white text on a black background
    mask = cv.inRange(hsv, lower_value, upper_value)
    # kernel = cv.getStructuringElement(cv.MORPH_RECT, (3, 3))
    # opening = cv.morphologyEx(mask, cv.MORPH_OPEN, kernel, iterations=1)  # Gets rid of small dots
    # Inverts the image to black text on white background
    invert = 255 - mask
    # Adds gaps in between characters so that they can be more easily recognized
    processed_image = add_space_between_characters(invert, 5)
    cv.imshow("A", processed_image)
    return processed_image


def crop_image(screenshot, rectangle):
    # rectangle = [location[0], location[1], image.shape[1], image.shape[0]]
    return screenshot[rectangle[0] + rectangle[2]:rectangle[1] + rectangle[3]]


def add_space_between_characters(cropped_screenshot: Image, gap: int) -> Image:
    """
    Adds a gap in between characters so that they can be read more easily
    :param cropped_screenshot: The image with the characters, the characters are black, the background is white
    :param gap: The amount of pixels added in between the characters
    :return: The processed image with added gaps
    """
    columns_to_be_added = [0]  # List storing the indexes of each gap to be added
    # Iterates through the image by column then row
    for x in range(2, len(cropped_screenshot[0]) - 2):  # Each column
        column_black_pixels = 0
        last_column_black_pixels = 0
        next_column_black_pixels = 0
        for y in range(len(cropped_screenshot)):  # Each row
            color = cropped_screenshot[y][x]
            last_color = cropped_screenshot[y][x - 2]
            next_color = cropped_screenshot[y][x + 2]
            # Checks if it is black
            if color < 50:
                column_black_pixels += 1
                # Makes the pixel completely black
                cropped_screenshot[y, x] = 0
            else:
                # Makes the pixel completely white
                cropped_screenshot[y, x] = 255
            if last_color < 50:
                last_column_black_pixels += 1
            if next_color < 50:
                next_column_black_pixels += 1
        # Gets the percentage of black pixels in the column
        percentage_of_black_pixels = column_black_pixels / cropped_screenshot.shape[0]
        percentage_of_last_black_pixels = last_column_black_pixels / cropped_screenshot.shape[0]
        percentage_of_next_black_pixels = next_column_black_pixels / cropped_screenshot.shape[0]
        # Checks if a gap should be added
        if percentage_of_last_black_pixels > .3 and percentage_of_black_pixels < .22 and percentage_of_next_black_pixels > .3:
            if (x - columns_to_be_added[-1]) > 3:
                columns_to_be_added.append(x)
            else:
                columns_to_be_added[-1] = x

    # Adds the gaps
    for column in columns_to_be_added[::-1]:
        for i in range(gap):
            cropped_screenshot = np.insert(cropped_screenshot, column, [255] * len(cropped_screenshot), axis=1)

    return cropped_screenshot


VK_CODE = {'backspace': 0x08,
           'tab': 0x09,
           'clear': 0x0C,
           'enter': 0x0D,
           'shift': 0x10,
           'ctrl': 0x11,
           'alt': 0x12,
           'pause': 0x13,
           'caps_lock': 0x14,
           'esc': 0x1B,
           'spacebar': 0x20,
           'page_up': 0x21,
           'page_down': 0x22,
           'end': 0x23,
           'home': 0x24,
           'left_arrow': 0x25,
           'up_arrow': 0x26,
           'right_arrow': 0x27,
           'down_arrow': 0x28,
           'select': 0x29,
           'print': 0x2A,
           'execute': 0x2B,
           'print_screen': 0x2C,
           'ins': 0x2D,
           'del': 0x2E,
           'help': 0x2F,
           '0': 0x30,
           '1': 0x31,
           '2': 0x32,
           '3': 0x33,
           '4': 0x34,
           '5': 0x35,
           '6': 0x36,
           '7': 0x37,
           '8': 0x38,
           '9': 0x39,
           'a': 0x41,
           'b': 0x42,
           'c': 0x43,
           'd': 0x44,
           'e': 0x45,
           'f': 0x46,
           'g': 0x47,
           'h': 0x48,
           'i': 0x49,
           'j': 0x4A,
           'k': 0x4B,
           'l': 0x4C,
           'm': 0x4D,
           'n': 0x4E,
           'o': 0x4F,
           'p': 0x50,
           'q': 0x51,
           'r': 0x52,
           's': 0x53,
           't': 0x54,
           'u': 0x55,
           'v': 0x56,
           'w': 0x57,
           'x': 0x58,
           'y': 0x59,
           'z': 0x5A,
           'numpad_0': 0x60,
           'numpad_1': 0x61,
           'numpad_2': 0x62,
           'numpad_3': 0x63,
           'numpad_4': 0x64,
           'numpad_5': 0x65,
           'numpad_6': 0x66,
           'numpad_7': 0x67,
           'numpad_8': 0x68,
           'numpad_9': 0x69,
           'multiply_key': 0x6A,
           'add_key': 0x6B,
           'separator_key': 0x6C,
           'subtract_key': 0x6D,
           'decimal_key': 0x6E,
           'divide_key': 0x6F,
           'F1': 0x70,
           'F2': 0x71,
           'F3': 0x72,
           'F4': 0x73,
           'F5': 0x74,
           'F6': 0x75,
           'F7': 0x76,
           'F8': 0x77,
           'F9': 0x78,
           'F10': 0x79,
           'F11': 0x7A,
           'F12': 0x7B,
           'F13': 0x7C,
           'F14': 0x7D,
           'F15': 0x7E,
           'F16': 0x7F,
           'F17': 0x80,
           'F18': 0x81,
           'F19': 0x82,
           'F20': 0x83,
           'F21': 0x84,
           'F22': 0x85,
           'F23': 0x86,
           'F24': 0x87,
           'num_lock': 0x90,
           'scroll_lock': 0x91,
           'left_shift': 0xA0,
           'right_shift ': 0xA1,
           'left_control': 0xA2,
           'right_control': 0xA3,
           'left_menu': 0xA4,
           'right_menu': 0xA5,
           'browser_back': 0xA6,
           'browser_forward': 0xA7,
           'browser_refresh': 0xA8,
           'browser_stop': 0xA9,
           'browser_search': 0xAA,
           'browser_favorites': 0xAB,
           'browser_start_and_home': 0xAC,
           'volume_mute': 0xAD,
           'volume_Down': 0xAE,
           'volume_up': 0xAF,
           'next_track': 0xB0,
           'previous_track': 0xB1,
           'stop_media': 0xB2,
           'play/pause_media': 0xB3,
           'start_mail': 0xB4,
           'select_media': 0xB5,
           'start_application_1': 0xB6,
           'start_application_2': 0xB7,
           'attn_key': 0xF6,
           'crsel_key': 0xF7,
           'exsel_key': 0xF8,
           'play_key': 0xFA,
           'zoom_key': 0xFB,
           'clear_key': 0xFE,
           '+': 0xBB,
           ',': 0xBC,
           '-': 0xBD,
           '.': 0xBE,
           '/': 0xBF,
           ';': 0xBA,
           '[': 0xDB,
           '\\': 0xDC,
           ']': 0xDD,
           "'": 0xDE,
           '`': 0xC0}
