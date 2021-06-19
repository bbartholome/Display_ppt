
# importing time and vlc
import time, vlc
import win32gui, win32con
# method to play video
def video(source):
    # creating a vlc instance
    vlc_instance = vlc.Instance()

    # creating a media player
    player = vlc_instance.media_player_new()

    # creating a media
    media = vlc_instance.media_new(source)

    # setting media to the player
    player.set_media(media)



    # play the video

    player.play()
    time.sleep(0.2)
  #  player.pause()
  #  time.sleep(0.5)
  #  player.play()
    # wait time

    # focus and maximize the vlc player
    handle = win32gui.FindWindow(None, "VLC (Direct3D11 output)")
    win32gui.SetForegroundWindow(handle)
    win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
    #getting the duration of the video
    duration = player.get_length()

    # printing the duration of the video
    print("Duration : " + str(duration))
    player.play()



    # wait time
    time.sleep(duration/1000)


# call the video method
video(r'\pps\videoplayback.mp4')
