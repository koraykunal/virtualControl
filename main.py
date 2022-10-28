import cv2
import numpy as np
import HandTrackingModule as htm
import time
import autopy
import pyautogui
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

######################
wCam, hCam = 640, 480
frameR = 100     #Frame Reduction
smoothening = 7  #random value
######################

pTime = 0
plocX, plocY = 0, 0
clocX, clocY = 0, 0
cap = cv2.VideoCapture(1)
cap.set(3, wCam)
cap.set(4, hCam)

detector = htm.handDetector(maxHands=1)
wScr, hScr = autopy.screen.size()

destroyMode = False

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
vol = 0
volBar = 400
volPer = 0
area = 0
colorVol = (255, 0, 0)
isClicked = False

while True:
    # Find the landmarks
    success, img = cap.read()
    img = detector.findHands(img)
    lmList, bbox = detector.findPosition(img)


    # Get the tip of the index and middle finger
    if len(lmList) != 0:
        x1, y1 = lmList[8][1:]
        x2, y2 = lmList[12][1:]
        x3, y3 = lmList[4][1], lmList[4][2]
        x4, y4 = lmList[8][1], lmList[8][2]
        #print(lmList)

        # Check which fingers are up
        fingers = detector.fingersUp()
        print(fingers)
        cv2.rectangle(img, (frameR, frameR), (wCam - frameR, hCam - frameR),
                      (255, 0, 255), 2)

        if fingers[3] == 0 and fingers[0] == 1 and fingers[4] == 1:
            area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]) // 100
            if 250 < area < 1000:

                # Find Distance between index and Thumb
                length, img, lineInfo = detector.findDistance(4, 8, img)

                # Convert Volume
                volBar = np.interp(length, [50, 200], [400, 150])
                volPer = np.interp(length, [50, 200], [0, 100])

                # Reduce Resolution to make it smoother
                smoothness = 10
                volPer = smoothness * round(volPer / smoothness)

                # Check fingers up
                fingers = detector.fingersUp()

                # If pinky is down set volume
                if fingers[4]:
                    volume.SetMasterVolumeLevelScalar(volPer / 100, None)
                    cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                    colorVol = (0, 255, 0)

                    cv2.rectangle(img, (50, 150), (85, 400), (255, 0, 0), 3)
                    cv2.rectangle(img, (50, int(volBar)), (85, 400), (255, 0, 0), cv2.FILLED)
                    cv2.putText(img, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX,
                                1, (255, 0, 0), 3)

            cv2.rectangle(img, (50, 150), (85, 400), (255, 0, 0), 3)
            cv2.rectangle(img, (50, int(volBar)), (85, 400), (255, 0, 0), cv2.FILLED)
            cv2.putText(img, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX,
                        1, (255, 0, 0), 3)

        # Only Index Finger: Moving Mode
        if fingers[1] == 1 and fingers[2] == 0 and fingers[0] == 1:

            # Convert the coordinates
            x3 = np.interp(x1, (frameR, wCam-frameR), (0, wScr))
            y3 = np.interp(y1, (frameR, hCam-frameR), (0, hScr))

            # Smooth Values
            clocX = plocX + (x3 - plocX) / smoothening
            clocY = plocY + (y3 - plocY) / smoothening

            # Move Mouse
            autopy.mouse.move(wScr - clocX, clocY)
            cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
            plocX, plocY = clocX, clocY

        # Both Index and middle are up: Clicking Mode
        if fingers[1] == 1 and fingers[2] == 1:
            # Find distance between fingers
            length, img, lineInfo = detector.findDistance(8, 12, img)

        # If raised up only pinky finger, scroll up ( usually effective on websites)
        if fingers[0] == 0 and fingers[4] == 1:
            pyautogui.scroll(100)

        # If raised up only thumb finger, scroll down
        elif fingers[0] == 1 and fingers[1] == 0:
            pyautogui.scroll(-100)

        # If you put your thumb down when your index finger up, click
        if fingers[0] == 0 and fingers[2] == 0 and isClicked == False:
            autopy.mouse.click()
            isClicked = True

        # If you put your thumb down when your index and middle finger up, double click
        if fingers[0] == 0 and fingers[2] == 1 and isClicked == False:
            pyautogui.doubleClick()
            isClicked = True

        # If you put your thumb down when your index, middle and ring finger up, right click
        if fingers[0] == 0 and fingers[2] == 1 and fingers[3] == 1:
            pyautogui.click(button='right', clicks=1, interval=0.15)

        # When return moving mode, change clicked variable to false
        if fingers[1] == 1 and fingers[0] == 1:
            isClicked = False

        # When all fingers up, enable destroy mode
        if fingers[1] == 1 and fingers[2] == 1 and fingers[3] == 1 and fingers[4] == 1 and fingers[0] == 1:
            destroyMode = True

        # And during this time if you put all fingers down close current window
        if fingers[1] == 0 and fingers[2] == 0 and fingers[3] == 0 and fingers[4] == 0 and fingers[0] == 0 and destroyMode == True:
            pyautogui.hotkey('altleft', 'f4')
            destroyMode = False

    # Frame rate
    cTime = time.time()
    fps = 1/(cTime-pTime)
    pTime = cTime
    cv2.putText(img, str(int(fps)), (28, 58), cv2.FONT_HERSHEY_PLAIN, 3, (255, 8, 8), 3)

    # Display
    cv2.imshow("Image", img)
    cv2.waitKey(1)