import win32gui
import win32api
import random
import time
import win32con

def MoveAndClick(left, top, hwnd):#鼠標在目標窗口上的點擊操作函數（點擊的位置進行了隨機，點擊間隔時間也進行了隨機）

    click_x = left + 460#點擊點的X座標
    click_y = top + 540#點擊點的Y座標
    #print('click_x=',click_x,'click_y=',click_y)  #檢查點擊點的座標
    random_deltax = (random.random() - 0.5) * 15#在點擊點的附近，x座標上產生一個[-7.5,7.5]區間的均勻分佈的隨機波動
    random_deltay = (random.random() - 0.5) * 15#在點擊點的附近，y座標上產生一個[-7.5,7.5]區間的均勻分佈的隨機波動
    click_pos =win32gui.ScreenToClient(hwnd,(click_x,click_y))#將點擊點左邊變成需要點擊的界面上的座標（這一步必須要有！）
    tmp = win32api.MAKELONG(click_pos[0], click_pos[1])#將座標轉換成SendMessage()函數需要的一維數據的座標
    #print('tmp = ',tmp)
    win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE,win32con.WA_ACTIVE,0)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN,win32con.MK_LBUTTON,tmp)#鼠標左鍵在tmp位置按下
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP,0000,tmp)#鼠標左鍵在tmp位置彈起
    #（彈起時，鼠標左鍵的代碼是0000，有些電腦可能是win32con.MK_LBUTTON,具體情況需要通過手動在界面上點擊鼠標左鍵，看spy++中的顯示）
    pause_t = random.random() * 3#產生一個隨機0-3s的延時
    count = random.random() * 10#產生一個隨機0-10的計數循環
    #print('int(count)=',int(count))
    for i in range(1000):
        time.sleep(pause_t)
        for i in range(int(count)):#隨機循環點擊0-10次，每次間隔時間是0-1s
            random_deltax = (random.random() - 0.5) * 15#在點擊點的附近，x座標上產生一個[-7.5,7.5]區間的均勻分佈的隨機波動
            random_deltay = (random.random() - 0.5) * 15#在點擊點的附近，y座標上產生一個[-7.5,7.5]區間的均勻分佈的隨機波動
            pause_random_t = random.random()#產生一個隨機0-1s的延時
            time.sleep(pause_random_t)#實現延時
            #將點擊點左邊變成需要點擊的界面上的座標（這一步必須要有！）
            click_pos =win32gui.ScreenToClient(hwnd,(click_x+int(random_deltax),click_y+int(random_deltay)))
            tmp = win32api.MAKELONG(click_pos[0], click_pos[1])#將座標轉換成SendMessage()函數需要的一維數據的座標
            #print('tmp = ',tmp)
            win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE,win32con.WA_ACTIVE,0)
            win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN,win32con.MK_LBUTTON,tmp)#鼠標左鍵在tmp位置按下
            #time.sleep(t)#延遲時間t秒以後再彈起
            win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP,0000,tmp)#鼠標左鍵在tmp位置彈起

MoveAndClick()