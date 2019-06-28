import win32gui
import win32ui
from ctypes import windll
from PIL import Image
import os
import keyboard

top_hwnd = None

def bringtofront(windowname=None):
    HWND = win32gui.FindWindow(None, windowname)
    win32gui.SetForegroundWindow(HWND)

def callback(hwnd, strings):
    if win32gui.IsWindowVisible(hwnd):
        window_title = win32gui.GetWindowText(hwnd)
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        if window_title and right-left and bottom-top:
            strings.append('0x{:08x}: "{}"'.format(hwnd, window_title))
    return True

def findOutlook():
    global top_hwnd
    win_list = []  # list of strings containing win handles and window titles
    win32gui.EnumWindows(callback, win_list)  # populate list

    for window in win_list:  # print results
        if "Outlook" in window:
            windowPos = win_list.index(window)
            #print(win_list.index(window))
            top_hwnd = win_list[windowPos][13:-1] #slice shit off
            bringtofront(top_hwnd)
            keyboard.press_and_release("ctrl + 2")
            keyboard.press_and_release("ctrl + alt + 2")
            keyboard.press_and_release("alt + h, o + d")
            keyboard.press_and_release("alt + down")
            keyboard.press_and_release("alt + h, o + d")
            keyboard.press_and_release("ctrl + 1")

def next_week_planned():
    findOutlook()
    try:
        hwnd = win32gui.FindWindow(None, top_hwnd)
        # Change the line below depending on whether you want the whole window
        # or just the client area.
        #left, top, right, bot = win32gui.GetClientRect(hwnd)
        left, top, right, bot = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bot - top

        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

        saveDC.SelectObject(saveBitMap)

        # Change the line below depending on whether you want the whole window
        # or just the client area.
        #result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)
        #print(result)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

        if result == 1:
            #PrintWindow Succeeded
            im.save("./test.jpg")
            imageobject = Image.open("./test.jpg")
            cropped = imageobject.crop((245,170,1920,730))
            cropped = cropped.resize((int(cropped.width*.75),int(cropped.height * .75)), Image.ANTIALIAS)
            cropped.save("./cropped.jpg")
            os.remove("./test.jpg")
    except Exception as e:
        print("\n" + str(e))

