import cv2
import HandTrackingModule as htm
import time
import win32com.client
import win32api
import win32con

screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)  # 获取屏幕宽度
screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)  # 获取屏幕高度

print("屏幕宽度：", screen_width)
print("屏幕高度：", screen_height)
x0=0
y0=0
num0=0
num1=0
num2=0
flag0=False
flag1=False
flag2=False
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True
presentation = Application.Presentations.Open("D:\\python3\\python.pptx")

cap = cv2.VideoCapture(0)
cap.set(4, 1440)
cap.set(3, 2560)
pTime = 0
detector = htm.handDetector()

while True:
    success, img = cap.read()
    img = detector.findHands(img)
    lmList = detector.findPosition(img, draw=False)
    pointList = [4, 8, 12, 16, 20]
    if len(lmList) != 0:
        countList = []

        if lmList[4][1] < lmList[3][1]:
            countList.append(1)
        else:
            countList.append(0)

        for i in range(1, 5):
            if lmList[pointList[i]][2] < lmList[pointList[i] - 2][2]:
                countList.append(1)
            else:
                countList.append(0)

        count = countList.count(1)
        count = int(count)
        cv2.putText(img, f'{count}', (15, 400), cv2.FONT_HERSHEY_PLAIN, 15, (255, 0, 255), 10)

        cv2.imshow("Image", img)
        if count == 5:
            settings = presentation.SlideShowSettings.Run()

            cv2.waitKey(1)
        if lmList[8][2] > lmList[4][2]:
            if lmList[pointList[1]][1] > lmList[pointList[1] - 2][1]:
                if flag2 == True:
                    num2 += 1
                    cv2.waitKey(1)
                else:
                    flag2 = True
                    num2=0
                if num2 > 5:
                    num2=0
                    presentation.SlideShowWindow.View.Previous()
                    cv2.waitKey(500)
                flag1,flag0=False,False
            elif lmList[pointList[1]][1] < lmList[pointList[1] - 2][1] :
                if flag1 == True:
                    num1 += 1
                    cv2.waitKey(1)
                else:
                    flag1 = True
                    num1=0
                if num1 > 5:
                    presentation.SlideShowWindow.View.Next()
                    cv2.waitKey(500)
                    num1=0
                flag0,flag2=False,False
            elif count == 0:
                flag1,flag2=False,False
        elif lmList[pointList[1]][2] > lmList[pointList[2]][2]:
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            cv2.waitKey(500)
        elif lmList[pointList[1]][2] < lmList[pointList[1] - 2][2]:
            if x0==0 and y0==0:
                x0=lmList[pointList[1]][1]
                y0=lmList[pointList[1]][2]

            x1 = lmList[pointList[1]][1]-x0
            y1 = lmList[pointList[1]][2]-y0
            x0 = lmList[pointList[1]][1]
            if abs(x1)>2 or abs(y1)>2:
                x, y = win32api.GetCursorPos()
                win32api.SetCursorPos((x-x1*2, y+y1*2))
            y0 = lmList[pointList[1]][2]
            cv2.waitKey(1)





