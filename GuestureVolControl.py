import time
import cv2 as cv
import numpy as np
import handTrackingModule as htm
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

####################################
wcam, hcam = 640, 480
####################################

cap = cv.VideoCapture(0)
cap.set(3, wcam)
cap.set(4, hcam)

pTime = 0
detector = htm.handDetector(detectionCon=0.7)

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
# volume.GetMasterVolumeLevel()

volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
vol = 0
volBar = 0
volPer = 0

volume.SetMasterVolumeLevel(-20.0, None)
while True:
    success, img = cap.read()
    img = detector.findHands( img )
    lmList = detector.findPosition( img )
    if len( lmList ) != 0:
        # print( lmList[4] )
        x1, y1 = lmList[4][1], lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]
        cx, cy = (x1+x2)//2, (y1+y2)//2
        cv.circle(img, (x1, y1), 10, (0,255,0),-1)
        cv.circle( img, (x2, y2), 10, (0,255, 0), -1 )
        cv.line(img, (x1, y1), (x2, y2), (128, 128, 128), 3)
        cv.circle(img, (cx, cy), 13, (0, 255, 0), -1)
        length = math.hypot(x2-x1, y2-y1)
        # print(length)

        # hand Range = 50 - 200
        # Volume Range = -95, 0

        vol = np.interp(length, [50,200], [minVol, maxVol])
        volBar = np.interp( length, [50, 200], [400, 150] )
        volPer = np.interp( length, [50, 200], [0, 200] )
        print(vol)
        volume.SetMasterVolumeLevel( vol, None )

        if length<50:
            cv.circle( img, (cx, cy), 13, (0, 255, 255), -1 )

    cv.rectangle(img, (50, 150), (85, 400), (255, 0, 0), 3)
    cv.rectangle(img, (50, int(volBar)), (85, 400), (255, 0, 0), -1)
    cv.putText( img, f"{int(volPer/2)}%", (40, 450), cv.FONT_HERSHEY_COMPLEX, 1,
                (0,  255, 0), 3 )





    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv.putText( img, str( int( fps ) ), (10, 70), cv.FONT_HERSHEY_PLAIN, 3,
                 (255, 0, 255), 3 )
    cv.imshow( "Image", img )

    if cv.waitKey( 1 ) & 0xFF == ord( "d" ):
        break
