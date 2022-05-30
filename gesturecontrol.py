import cv2
import numpy as np
from HandTrackingModule import HandDetector  # 手不检测方法
import time
import autopy
import win32gui, win32process, psutil
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volumeRange = volume.GetVolumeRange()  # (-63.5, 0.0, 0.03125)
minVol = volumeRange[0]
maxVol = volumeRange[1]

# （1）导数视频数据
wScr, hScr = autopy.screen.size()  # 返回电脑屏幕的宽和高(1920.0, 1080.0)
wCam, hCam = 1280, 720  # 视频显示窗口的宽和高
pt1, pt2 = (100, 100), (1000, 500)  # 虚拟鼠标的移动范围，左上坐标pt1，右下坐标pt2

cap = cv2.VideoCapture(0)  # 0代表自己电脑的摄像头
cap.set(3, wCam)  # 设置显示框的宽度1280
cap.set(4, hCam)  # 设置显示框的高度720

pTime = 0  # 设置第一帧开始处理的起始时间
pLocx, pLocy = 0, 0  # 上一帧时的鼠标所在位置
smooth = 5  # 自定义平滑系数，让鼠标移动平缓一些
frame = 0  # 初始化累计帧数
toggle = False  # 标志变量
prev_state = [1, 1, 1, 1, 1]  # 初始化上一帧状态
current_state = [1, 1, 1, 1, 1]  # 初始化当前正状态

# （2）接收手部检测方法
detector = HandDetector(mode=False,  # 视频流图像
                        maxHands=1,  # 最多检测一只手
                        detectionCon=0.8,  # 最小检测置信度
                        minTrackCon=0.5)  # 最小跟踪置信度

# （3）处理每一帧图像
while True:
    # 图片是否成功接收、img帧图像
    success, img = cap.read()
    # 翻转图像，使自身和摄像头中的自己呈镜像关系
    img = cv2.flip(img, flipCode=1)  # 1代表水平翻转，0代表竖直翻转
    # 在图像窗口上创建一个矩形框，在该区域内移动鼠标
    cv2.rectangle(img, pt1, pt2, (0, 255, 255), 5)
    # 判断当前的活动窗口的进程名字
    try:
        pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
        print("pid:", pid)
        active_window_process_name = psutil.Process(pid[-1]).name()
        print("acitiveprocess:", active_window_process_name)
    except:
        pass
    # （4）手部关键点检测
    # 传入每帧图像, 返回手部关键点的坐标信息(字典)，绘制关键点后的图像
    hands, img = detector.findHands(img, flipType=False, draw=True)  # 上面反转过了，这里就不用再翻转了
    print("hands:", hands)
    # [{'lmList': [[889, 652, 0], [807, 613, -25], [753, 538, -39], [723, 475, -52], [684, 431, -66], [789, 432, -27],
    #              [762, 347, -56], [744, 295, -78], [727, 248, -95], [841, 426, -39], [835, 326, -65], [828, 260, -89],
    #              [820, 204, -106], [889, 445, -54], [894, 356, -85], [892, 295, -107], [889, 239, -123],
    #              [933, 483, -71], [957, 421, -101], [973, 376, -115], [986, 334, -124]], 'bbox': (684, 204, 302, 448),
    #   'center': (835, 428), 'type': 'Right'}]
    # 如果能检测到手那么就进行下一步
    if hands:

        # 获取手部信息hands中的21个关键点信息
        lmList = hands[0]['lmList']  # hands是由N个字典组成的列表，字典包括每只手的关键点信息,此处代表第0个手
        hand_center = hands[0]['center']
        drag_flag = 0
        # 获取食指指尖坐标，和中指指尖坐标
        x1, y1, z1 = lmList[8]  # 食指尖的关键点索引号为8
        x2, y2, z2 = lmList[12]  # 中指指尖索引12
        cx, cy, cz = (x1 + x2) // 2, (y1 + y2) // 2, (z1 + z2) // 2  # 计算食指和中指两指之间的中点坐标
        hand_cx, hand_cy = hand_center[0], hand_center[1]
        # （5）检查哪个手指是朝上的
        fingers = detector.fingersUp(hands[0])  # 传入
        print("fingers", fingers)  # 返回 [0,1,1,0,0] 代表 只有食指和中指竖起
        # 255, 0,255 淡紫
        # 0,255,255  淡蓝
        # 255,255,0  淡黄
        # 计算食指尖和中指尖之间的距离distance,绘制好了的图像img,指尖连线的信息info
        distance, info, img = detector.findDistance((x1, y1), (x2, y2), img)  # 会画圈
        # （6）确定鼠标移动的范围
        # 将食指指尖的移动范围从预制的窗口范围，映射到电脑屏幕范围
        x3 = np.interp(x1, (pt1[0], pt2[0]), (0, wScr))
        y3 = np.interp(y1, (pt1[1], pt2[1]), (0, hScr))
        # 手心坐标映射到屏幕范围
        x4 = np.interp(hand_cx, (pt1[0], pt2[0]), (0, wScr))
        y4 = np.interp(hand_cy, (pt1[1], pt2[1]), (0, hScr))
        # （7）平滑，使手指在移动鼠标时，鼠标箭头不会一直晃动
        cLocx = pLocx + (x3 - pLocx) / smooth  # 当前的鼠标所在位置坐标
        cLocy = pLocy + (y3 - pLocy) / smooth
        # 记录当前手势状态
        current_state = fingers
        # 记录相同状态的帧数
        if (prev_state == current_state):
            frame = frame + 1
        else:
            frame = 0
        prev_state = current_state

        if fingers != [0, 0, 0, 0, 0] and toggle and frame >= 2:
            autopy.mouse.toggle(None, False)
            toggle = False
            print("释放左键")

        # 只有食指和中指竖起，就认为是移动鼠标
        if fingers[1] == 1 and fingers[2] == 1 and sum(fingers) == 2 and frame >= 1:
            # （8）移动鼠标
            autopy.mouse.move(cLocx, cLocy)  # 给出鼠标移动位置坐标

            print("移动鼠标")

            # 更新前一帧的鼠标所在位置坐标，将当前帧鼠标所在位置，变成下一帧的鼠标前一帧所在位置
            pLocx, pLocy = cLocx, cLocy

            # （9）如果食指和中指都竖起，指尖距离小于某个值认为是单击鼠标
            # 当指间距离小于43（像素距离）就认为是点击鼠标
            if distance < 43 and frame >= 1:
                # 在食指尖画个绿色的圆，表示点击鼠标
                cv2.circle(img, (x1, y1), 15, (0, 255, 0), cv2.FILLED)

                # 左击鼠标
                autopy.mouse.click(button=autopy.mouse.Button.LEFT, delay=0)
                cv2.putText(img, "left_click", (150, 50), cv2.FONT_HERSHEY_PLAIN, 3, (0, 255, 0), 3)
                print("左击鼠标")
            else:
                cv2.putText(img, "move", (150, 50), cv2.FONT_HERSHEY_PLAIN, 3, (0, 255, 0), 3)
        # 中指弯下食指在上，右击鼠标
        elif fingers[1] == 1 and fingers[2] == 0 and sum(fingers) == 1 and frame >= 2:
            autopy.mouse.click(button=autopy.mouse.Button.RIGHT, delay=0)
            print("右击鼠标")
            cv2.putText(img, "rigth_click", (150, 50), cv2.FONT_HERSHEY_PLAIN, 3, (0, 255, 0), 3)
            cv2.circle(img, (x2, y2), 15, (0, 255, 0), cv2.FILLED)

        # 五指紧握，按紧左键进行拖拽
        elif fingers == [0, 0, 0, 0, 0]:
            if toggle == False:
                autopy.mouse.toggle(None, True)
                print("按紧左键")
            toggle = True
            autopy.mouse.move(cLocx, cLocy)
            pLocx, pLocy = cLocx, cLocy
            cv2.putText(img, "drag", (150, 50), cv2.FONT_HERSHEY_PLAIN, 3, (0, 255, 0), 3)
            print("拖拽鼠标")

        # 拇指张开，其他弯曲，按一次上键
        elif fingers == [1, 0, 0, 0, 0] and frame >= 2:
            cv2.putText(img, "UP", (150, 50), cv2.FONT_HERSHEY_PLAIN, 3, (0, 255, 0), 3)
            if (active_window_process_name == "cloudmusic.exe"):
                print("#############################################")
                autopy.key.toggle(autopy.key.Code.LEFT_ARROW, True, [autopy.key.Modifier.CONTROL])
                autopy.key.toggle(autopy.key.Code.LEFT_ARROW, False, [autopy.key.Modifier.CONTROL])
                print("上一曲")
                time.sleep(0.3)
            else:
                autopy.key.toggle(autopy.key.Code.UP_ARROW, True, [])
                autopy.key.toggle(autopy.key.Code.UP_ARROW, False, [])
                print("按下上键")

                time.sleep(0.3)

        # 拇指弯曲，其他竖直，按一次下键
        elif fingers == [0, 1, 1, 1, 1] and frame >= 2:
            cv2.putText(img, "Down", (150, 50), cv2.FONT_HERSHEY_PLAIN, 3, (0, 255, 0), 3)
            if (active_window_process_name == "cloudmusic.exe"):
                print("#############################################")
                autopy.key.toggle(autopy.key.Code.RIGHT_ARROW, True, [autopy.key.Modifier.CONTROL])
                autopy.key.toggle(autopy.key.Code.RIGHT_ARROW, False, [autopy.key.Modifier.CONTROL])
                print("下一曲")
                time.sleep(0.3)
            else:
                autopy.key.toggle(autopy.key.Code.DOWN_ARROW, True, [])
                autopy.key.toggle(autopy.key.Code.DOWN_ARROW, False, [])

                print("按下下键")
                time.sleep(0.3)

        # 类ok手势，进行调整音量
        elif fingers == [1, 0, 1, 1, 1] and frame >= 5:
            autopy.mouse.move(cLocx, cLocy)  # 给出鼠标移动位置坐标
            length = cLocx - pLocx
            pLocx = cLocx
            pLocy = cLocy
            print("移动的length:", length)
            print("移动鼠标调整音量")
            currentVolumeLv = volume.GetMasterVolumeLevelScalar()
            print("currentVolume:", currentVolumeLv)
            currentVolumeLv += length / 50.0
            if currentVolumeLv > 1.0:
                currentVolumeLv = 1.0
            elif currentVolumeLv < 0.0:
                currentVolumeLv = 0.0
            volume.SetMasterVolumeLevelScalar(currentVolumeLv, None)
            setVolume = volume.GetMasterVolumeLevelScalar()
            volPer = setVolume
            volBar = 350 - int((volPer) * 200)
            cv2.rectangle(img, (20, 150), (50, 350), (255, 0, 255), 2)
            cv2.rectangle(img, (20, int(volBar)), (50, 350), (255, 0, 255), cv2.FILLED)
            cv2.putText(img, f'{int(volPer * 100)}%', (10, 380), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 255), 2)

    # （10）显示图像
    # 查看FPS
    cTime = time.time()  # 处理完一帧图像的时间
    fps = 1 / (cTime - pTime)
    pTime = cTime  # 重置起始时·
    print(fps)
    # 在视频上显示fps信息，先转换成整数再变成字符串形式，文本显示坐标，文本字体，文本大小
    cv2.putText(img, str(int(fps)), (70, 50), cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 0), 3)

    # 显示图像，输入窗口名及图像数据
    cv2.imshow('frame', img)
    if cv2.waitKey(1) & 0xFF == 27:  # 每帧滞留20毫秒后消失，ESC键退出
        break

# 释放视频资源
cap.release()
cv2.destroyAllWindows()

# if __name__ == '__main__':
