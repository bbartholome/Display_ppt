
# importing time and vlc
import time, vlc
import win32gui, win32con



def video(source):
    # https://www.olivieraubert.net/vlc/python-ctypes/doc/vlc.Instance-class.html
    # creating a vlc instance
    vlc_instance = vlc.Instance("video")

    # creating a media player
    player = vlc_instance.media_player_new()

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

    #os.system("TASKKILL /F /IM vlc.exe")
    vlc_instance.vlm_del_media(source)


# call the video method

i=0
while i<5:
    i=+1
    print(i)
    video(r'\pps\videoplayback.mp4')
