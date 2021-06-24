# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

# https://python-pptx.readthedocs.io/en/latest/api/presentation.html
# https://docs.microsoft.com/en-us/office/vba/api/powerpoint.slideshowview.next

import time,shutil,os
from win32com.client import Dispatch   # pywin32 package
import configparser
from os import walk
import threading as th
import keyboard  #keyboard package
import vlc
import win32gui, win32con

keep_going = True

#*******************************************
def Clone_File(source, target):
    original = []  # get the list of file in directory 'path'
    for (dirpath, dirnames, original_filenames) in walk(source):
        original.extend(original_filenames)
        break
    print('file list:', original)

    copy = []  # get the list of file in directory 'path'
    for (dirpath, dirnames, copy_filenames) in walk(target):
        copy.extend(copy_filenames)
        break
    print('file list:', copy)

    for name in original:   #if file not presnt in destination folder copy it
        present=False
        for dest_name in copy:
            if name == dest_name:
                present=True
        if not present:
            try:
                shutil.copyfile(source + name, target + name)
                print("file copied", source + name, " to ", target + name )
            except:
                print('failed to copy')
        else:
            print("file already present", source + name )

    for dest_name in copy: #if file not present in source directory erase it in destination directory
        present = False
        for name in original:
            if name == dest_name:
                present = True
        if not present:
            try:
                os.remove(target + dest_name)
                print("file deleted from", target + dest_name)
            except FileNotFoundError :
                print('file not found, failed to delete',target + dest_name)


        else:
            print("file present", source + name)


#***********************************************
def key_capture_thread():
    global keep_going
    a = keyboard.read_key()
    if a == "esc":
        keep_going = False
        print('Le programme va arrêter')

#***********************************************
def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

#***********************************************
def get_app_list(handles=[]):
    mlst=[]
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst

#***********************************************
def Check_file_type(filename):
    if filename.find(".ppt") != -1 or filename.find(".pps") != -1 or filename.find(".PPT") != -1 or filename.find(".PPS") != -1:
       return "ppt"
    elif filename.find(".MOV") != -1 or filename.find(".mov") != -1:
        return "video"
    elif filename.find(".mp4") != -1 or filename.find(".MP4") != -1:
        return "video"

#***********************************************
def loop_file(path, delay):
    print(path)
    f = []  # get the list of file in directory 'path'
    for (dirpath, dirnames, filenames) in walk(path):
        f.extend(filenames)
        break
    print('file list:', f)
    for name in f:  # open and show the ppt
        if keep_going:

            type = Check_file_type(name)
            if type == "ppt":
                Read_ppt(PPT,path, name, delay)
            elif type == "video":
                video(path+name,VLC_instance,Player)


#***********************************************
def video(source, vlc_instance, player):
    #https://www.olivieraubert.net/vlc/python-ctypes/doc/vlc.Instance-class.html


    if vlc_instance==None:
        vlc_instance = vlc.Instance('video')   # creating a vlc instance

    if player==None:
        player = vlc_instance.media_player_new() # creating a media player

    # creating a media
    media = vlc_instance.media_new(source)

    # setting media to the player
    player.set_media(media)

    # play the video

    player.play()
    # media_
    player.toggle_fullscreen()
    time.sleep(0.2)
    player.pause()

    # wait time
    time.sleep(1)
    player.play()
    # focus and maximize the vlc player
    try:
        handle = win32gui.FindWindow(None, "VLC (Direct3D11 output)")
        win32gui.SetForegroundWindow(handle)
        win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
    except:
        print('windows not found')

    # getting the duration of the video
    duration = player.get_length()

    # printing the duration of the video
    # print("Duration : " + str(duration))

    # wait video time time
    time.sleep(duration / 1000)

    vlc_instance.vlm_del_media('video')


#***********************************************
def Read_ppt(ppt,path,filename, delay):
    # Use a breakpoint in the code line below to debug your script.
    print(path+filename)

    if ppt==None:
        ppt = Dispatch('Powerpoint.Application')
        ppt.Visible = True  # optional: if you want to see the spreadsheet
        ppt.Activate

    print('filename:', filename)

    try:
        pptfile = ppt.Presentations.Open(path + filename, 1)  # open presentation (readOnly)
        print('slide count:', ppt.ActivePresentation.Slides.Count)
    except:
        print("error opening file", path+filename)

    time.sleep(2)
    appwindows = get_app_list()
    for i in appwindows:
        print(i)
        if i[1].find("PowerPoint") != -1 and i[1].find(filename) != -1:
            try:
                handle = win32gui.FindWindow(0, i[1])
                win32gui.SetForegroundWindow(handle)
                win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
            except:
                print('error sw maximize')

    j = 0
    while (j < ppt.ActivePresentation.Slides.Count) and keep_going:
        if j==0 and filename.find(".ppt") != -1:
            ppt.ActivePresentation.SlideShowSettings.Run()  # needed if PPTX and ppt file not needed if PPSX and pps file

        j += 1
        print('shape count', ppt.ActivePresentation.Slides(j).Shapes.Count)

        SleepTime = 0
        k = 0
        while k < ppt.ActivePresentation.Slides(j).Shapes.Count:
            k += 1
            print('forme :', k)
            print('forme type :', ppt.ActivePresentation.Slides(j).Shapes(k).Type)
            if ppt.ActivePresentation.Slides(j).Shapes(k).Type==16:  # value 16 meams Media  #https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetypeforme
                VideoLength = ppt.ActivePresentation.Slides(j).Shapes(k).MediaFormat.Length  # duration of the video in ms
                print('video length' , VideoLength / 1000)
                SleepTime += VideoLength/1000

        time.sleep(max(SleepTime,delay))
        ppt.SlideShowWindows(1).View.Next()
        #ppt.SlideShowWindows(1).View.Exit()
    pptfile.Close()

    if not(keep_going):
        ppt.Quit()


#********************************************

if __name__ == '__main__':

    # creating a vlc instance
    VLC_instance = vlc.Instance('video')

    # creating a media player
    Player = VLC_instance.media_player_new()

    PPT = Dispatch('Powerpoint.Application')
    PPT.Visible = True  # optional: if you want to see the spreadsheet
    PPT.Activate

    config = configparser.ConfigParser()
    try:
        config.read('PowerPoint.ini')  # read config file
    except:
        print("cannot open ini file")

    path = config.get('Path', 'Path')  # get the path for the ppt file
    Drop_Path= config.get('Path', 'Drop_File_Path')  # get the path for the ppt file
    delay = config.getint('Delay', 'Delay')  # get the stop time between 2 slides

    th.Thread(target=key_capture_thread, args=(), name='key_capture_thread', daemon=True).start()
    while keep_going:
        print('still going...')
        print('check directory')
        Clone_File(Drop_Path,path)
        print('looping..')
        loop_file(path, delay)


    print('Fin de programme')


# See PyCharm help at https://www.jetbrains.com/help/pycharm/