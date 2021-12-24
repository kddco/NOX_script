import win32gui
import win32api
import random
import time
import win32con


# win32gui和win32api有大量函數是重複的，也就是說既能通過gui來調用，也能通過api來調用，比如下面的FindWindow(),FindWindowEx(),GetWindowText(),GetWindowRect(),GetCursorPos()函數

def GetXY():  # 獲得模擬器的窗口位置
    hwnd = win32gui.FindWindow('Qt5QWindowIcon',
                               '夜神模擬器')  # 是文件句柄，通過使用visual studio自帶的spy++獲得的。在工具欄中的 工具->spy++中,Qt5QWindowIcon是窗口類名，夜神模擬器 是窗口標題(窗口標題不一定是你窗口左上角顯示的標題)
    # print(hwnd)
    # 可操作窗口一般不是主窗口，一般是子窗口，子窗口必須使用FindWindowEx()函數來進行搜索
    # hwnd = win32gui.FindWindowEx(hwnd, 0, 'Qt5QWindowIcon', 'ScreenBoardClassWindow');#在窗口句柄爲hwnd的窗口中，（本例是 夜神模擬器），尋找子窗口，同樣是在spy++工具中看到的窗口信息。Qt5QWindowIcon是窗口類名，ScreenBoardClassWindow是窗口標題
    # hwnd = win32gui.FindWindowEx(hwnd, 0, 'Qt5QWindowIcon', 'QWidgetClassWindow');#在窗口句柄爲hwnd的窗口中，（本例是 ScreenBoardClassWindow），尋找子窗口，同樣是在spy++工具中看到的窗口信息。Qt5QWindowIcon是窗口類名，ScreenBoardClassWindow是窗口標題
    # print('hwnd=',hwnd)
    text = win32gui.GetWindowText(hwnd)  # 返回的是窗口的名字（不一定是窗口左上角顯示的名字）
    print('操作的窗口名爲：', text)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)  # (left,top)是左上角的座標，(right,bottom)是右下角的座標
    # win32gui.SetForegroundWindow(hwnd)
    return (left, top, hwnd)  # 返回模擬器的左上角(x,y)座標（即left,top），以及模擬器窗口的句柄
    # mouseX, mouseY = win32gui.GetCursorPos() # 返回當前鼠標位置，注意座標系統中左上方是(0, 0)
    # print('mouseX=',mouseX,'mouseY=',mouseY)


def MoveAndClick(left, top, hwnd):  # 鼠標在目標窗口上的點擊操作函數（點擊的位置進行了隨機，點擊間隔時間也進行了隨機）
    click_x = left + 460  # 點擊點的X座標
    click_y = top + 540  # 點擊點的Y座標
    # print('click_x=',click_x,'click_y=',click_y)  #檢查點擊點的座標
    random_deltax = (random.random() - 0.5) * 15  # 在點擊點的附近，x座標上產生一個[-7.5,7.5]區間的均勻分佈的隨機波動
    random_deltay = (random.random() - 0.5) * 15  # 在點擊點的附近，y座標上產生一個[-7.5,7.5]區間的均勻分佈的隨機波動
    click_pos = win32gui.ScreenToClient(hwnd, (click_x, click_y))  # 將點擊點左邊變成需要點擊的界面上的座標（這一步必須要有！）
    tmp = win32api.MAKELONG(click_pos[0], click_pos[1])  # 將座標轉換成SendMessage()函數需要的一維數據的座標
    # print('tmp = ',tmp)
    win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)  # 鼠標左鍵在tmp位置按下
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0000, tmp)  # 鼠標左鍵在tmp位置彈起
    # （彈起時，鼠標左鍵的代碼是0000，有些電腦可能是win32con.MK_LBUTTON,具體情況需要通過手動在界面上點擊鼠標左鍵，看spy++中的顯示）
    pause_t = random.random() * 3  # 產生一個隨機0-3s的延時
    count = random.random() * 10  # 產生一個隨機0-10的計數循環
    # print('int(count)=',int(count))
    for i in range(1000):
        time.sleep(pause_t)
        for i in range(int(count)):  # 隨機循環點擊0-10次，每次間隔時間是0-1s
            random_deltax = (random.random() - 0.5) * 15  # 在點擊點的附近，x座標上產生一個[-7.5,7.5]區間的均勻分佈的隨機波動
            random_deltay = (random.random() - 0.5) * 15  # 在點擊點的附近，y座標上產生一個[-7.5,7.5]區間的均勻分佈的隨機波動
            pause_random_t = random.random()  # 產生一個隨機0-1s的延時
            time.sleep(pause_random_t)  # 實現延時
            # 將點擊點左邊變成需要點擊的界面上的座標（這一步必須要有！）
            click_pos = win32gui.ScreenToClient(hwnd, (
            click_x + int(random_deltax), click_y + int(random_deltay)))  # 將點擊點左邊變成需要點擊的界面上的座標（這一步必須要有！）
            tmp = win32api.MAKELONG(click_pos[0], click_pos[1])  # 將座標轉換成SendMessage()函數需要的一維數據的座標
            # print('tmp = ',tmp)
            win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
            # win32api.SendMessage(hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, tmp)
            win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)  # 鼠標左鍵在tmp位置按下
            # time.sleep(t)#延遲時間t秒以後再彈起
            win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0000, tmp)  # 鼠標左鍵在tmp位置彈起
            print('我點！')


def main():
    print('程序開始運行！')
    left, top, hwnd = GetXY()
    # print('left=',left,'top=',top,'hwnd=',hwnd)
    MoveAndClick(left, top, hwnd)


# 此處運行程序
main()
