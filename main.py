# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

# https://python-pptx.readthedocs.io/en/latest/api/presentation.html
# https://docs.microsoft.com/en-us/office/vba/api/powerpoint.slideshowview.next

import time
from win32com.client import Dispatch   # pywin32 package
import configparser
from os import walk
import threading as th
import keyboard                         #keyboard package

keep_going = True

def key_capture_thread():
    global keep_going
    a = keyboard.read_key()
    if a == "esc":
        keep_going = False
        print('Le programme va arrêter')


def Read_ppt(path, delay):
    # Use a breakpoint in the code line below to debug your script.
    print(path)


    f = []  # get the list of file in directory 'path'
    for (dirpath, dirnames, filenames) in walk(path):
        f.extend(filenames)
        break
    print(f)

    ppt = Dispatch('Powerpoint.Application')
    ppt.Visible = True  # optional: if you want to see the spreadsheet
    ppt.Activate
    # win32gui.ShowWindow(ppt.hwnd, win32con.SW_SHOWNORMAL)

    for i in f:  # open and show the ppt
        if keep_going:
            filename = path + i
            print(filename)
            pptfile = ppt.Presentations.Open(filename, 1)  #open presentation (readOnly)

            if  filename.find(".ppt") != -1:
                ppt.ActivePresentation.SlideShowSettings.Run()   #needed if PPTX and ppt file not needed if PPSX and pps file

            time.sleep(delay)
            print(ppt.ActivePresentation.Slides.Count)

            j = 1
            while (j < ppt.ActivePresentation.Slides.Count) and keep_going:
                j += 1
                ppt.SlideShowWindows(1).View.Next()
                time.sleep(delay)

            #ppt.SlideShowWindows(1).View.Exit()
            pptfile.Close()

    if not(keep_going):
        ppt.Quit()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    config = configparser.ConfigParser()
    config.read('PowerPoint.ini')  # read config file

    path = config.get('Path', 'Path')  # get the path for the ppt file
    delay = config.getint('Delay', 'Delay')  # get the stop time between 2 slides

    th.Thread(target=key_capture_thread, args=(), name='key_capture_thread', daemon=True).start()
    while keep_going:
        print('still going...')
        Read_ppt(path,delay)

    print('Fin de programme')
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
