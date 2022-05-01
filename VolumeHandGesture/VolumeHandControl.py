"""
Created on Sunday 1 May 2022

@author: Bhupesh Kumar Dewangan
"""


# importing the all modules
import cv2 as cv
import time
import math
import numpy as np
import HandTrackingModule as HTM

from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

detector = HTM.handTracker(detectionCon=0.7)
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)

volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()


def main():
    minVol = volRange[0]
    maxVol = volRange[1]
    vol = 0
    volBar = 400
    volPer = 0

    # ---------------------
    wCam, hCam = 1080, 960
    # ---------------------

    cap = cv.VideoCapture(0)
    cap.set(3, wCam)
    cap.set(4, hCam)
    pTime = 0

    while True:
        success, img = cap.read()
        img = cv.flip(img, 1)
        img = detector.handsFinder(img)

        lmlist = detector.positionFinder(img, draw=False)

        if len(lmlist) != 0:

            #co-ordinates of index finger and thumb and making a cicle on both
            x1, y1 = lmlist[4][1], lmlist[4][2]
            x2, y2 = lmlist[8][1], lmlist[8][2]
            cx, cy = (x1+x2)//2, (y1+y2)//2

            cv.circle(img, (x1, y1), 10, (255, 0, 255), cv.FILLED)
            cv.circle(img, (x2, y2), 10, (255, 0, 255), cv.FILLED)

            cv.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
            cv.circle(img, (cx, cy), 10, (255, 0, 255), cv.FILLED)

            length = math.hypot(x2 - x1, y2 - y1)
            # print(length)

            # Hand Range from 0 - 300
            # Volume Range from -63.5 - 0

            vol = np.interp(length, [20, 250], [minVol, maxVol])
            volBar = np.interp(length, [20, 250], [400, 150])
            volPer = np.interp(length, [20, 250], [0, 100])
            print(int(length),"|", volRange, "|",minVol, "|", maxVol, "|", vol)
            volume.SetMasterVolumeLevel(vol, None)

            if length < 35:
                cv.circle(img, (cx, cy), 10, (0, 255, 0), cv.FILLED)

        cv.rectangle(img, (50, 150), (85, 400), (60, 149, 111), 2)
        cv.rectangle(img, (50, int(volBar)), (85, 400), (60, 149, 111), cv.FILLED)

        cv.putText(img,"Mute" if volPer == 0 else f'PER : {int(volPer)}', (40, 450), cv.FONT_HERSHEY_COMPLEX, 1, (106, 116, 218), 3)

        cTime = time.time()
        fps = 1/(cTime-pTime)
        pTime = cTime

        cv.putText(img, f'FPS : {int(fps)}', (40, 50), cv.FONT_HERSHEY_COMPLEX, 1, (182, 164, 22), 3)

        cv.imshow('Volume Control', img)
        key = cv.waitKey(1)

        if key == 13 or (cv.getWindowProperty('Volume Control', cv.WND_PROP_VISIBLE) < 1):
            print("!-------------PROGRAM TERMINATED !!-------------!")
            cap.release()
            cv.destroyAllWindows()
            break


if __name__ == "__main__":
    main()