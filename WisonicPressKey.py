#coding:utf8
from ctypes import *
import win32api,win32gui
import win32con
import time
#输入的是按键，不能作为文字输入，在函数使用时，每次最多可输入3个按键。超出3个，或没有输入，会退出
class WisonicKey():
    def __init__(self):
        self.keymap={"A":65,"B":66,"C":67,"D":68,"E":69,"F":70,"G":71,"H":72,"I":73,
                     "J":74,"K":75,"L":76,"M":77,"N":78,"O":79,"P":80,"Q":81,"R":82,
                     "S":83,"T":84,"U":85,"V":86,"W":87,"X":88,"Y":89,"Z":90,"0":48,
                     "1":49,"2":50,"3":51,"4":52,"5":53,"6":54,"7":55,"8":56,"9":57,
                     "F1":112,"F2":113,"F3":114,"F4":115,"F5":116,"F6":117,"F7":118,
                     "F8":119,"F9":120,"F10":121,"F11":122,"F12":123,"ENTER":13,"CTRL":17,"ALT":18,
                     "CAPS":20,"ESC":27,"SPACE":32,"PAGEUPP":33,"PAGEDOWN":34,"END":35,"HOME":36,
                     "LEFT":37,"UP":38,"RIGHT":39,"DOWN":40,"TAB":9,"BACKSPACE":8,
                     '0':48,'1':49,'2':50,'3':51,'4':52,'5':53,'6':54,'7':55,'8':56,'9':57}
    def PressKey(self,*arg):
        # 按键最多同时按下三个
        if len(arg)==0:
            return 0
        if len(arg)==1:
            key=self.keymap[arg[0].upper()]
            win32api.keybd_event(key,0,0,0)
            time.sleep(0.2)
            win32api.keybd_event(key,0,win32con.KEYEVENTF_KEYUP,0)
            return 0
        if len(arg)==2:
            key1=self.keymap[arg[0].upper()]
            key2=self.keymap[arg[1].upper()]
            win32api.keybd_event(key1,0,0,0)
            win32api.keybd_event(key2,0,0,0)
            time.sleep(0.2)
            win32api.keybd_event(key2,0,win32con.KEYEVENTF_KEYUP,0)
            win32api.keybd_event(key1,0,win32con.KEYEVENTF_KEYUP,0)
            time.sleep(0.2)
            return 0
        if len(arg)==3:
            key31=self.keymap[arg[0].upper()]
            key32=self.keymap[arg[1].upper()]
            key33=self.keymap[arg[2].upper()]
            win32api.keybd_event(key31,0,0,0)
            win32api.keybd_event(key32,0,0,0)
            win32api.keybd_event(key33,0,0,0)
            time.sleep(0.2)
            win32api.keybd_event(key33,0,win32con.KEYEVENTF_KEYUP,0)
            win32api.keybd_event(key32,0,win32con.KEYEVENTF_KEYUP,0)
            win32api.keybd_event(key31,0,win32con.KEYEVENTF_KEYUP,0)
            time.sleep(0.2)
            return 0
        else:
            return 0
    def Input(self,*arg):
        # 输入
        if len(arg)==0:
            return 0
        elif len(arg)==1:
            for word in list(arg[0]):
                if word==' ':
                    word='space'
                key=self.keymap[word.upper()]
                win32api.keybd_event(key,0,0,0)
                time.sleep(0.2)
                win32api.keybd_event(key,0,win32con.KEYEVENTF_KEYUP,0)
            return 0
    def MouseMoveTo(self,x,y):
        # 移动到坐标位置
        windll.user32.SetCursorPos(x,y)
    def MouseMoveAndClick(self,x,y):
        # 移动到坐标位置并点击
        windll.user32.SetCursorPos(x,y)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y)
        time.sleep(0.5)
    def MouseDownMove(self,x,y):
        # 按下后拖住移动
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y)
        time.sleep(0.05)
        windll.user32.SetCursorPos(x,y)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y)
        time.sleep(0.5)
    def Click(self):
        # 点击
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.5)
    def RightClick(self):
        # 右键点击
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
        time.sleep(0.5)
    def getw1Window(self,*arg):
        '''
        # 通过窗口类和窗口名获取主窗口的方法，并返回窗口坐标
        :param arg: 窗口title
        :return:获取到的窗口坐标
        '''
        hwnd=0
        if len(arg)==0:
            while(not hwnd):
                try:
                    hwnd=win32gui.FindWindowEx(None,None,"Qt5QWindowIcon","w1")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.5)
                except:
                    print("no w1")
                    time.sleep(2)
                    pass
        elif len(arg)==1:
            while(not hwnd):
                try:
                    hwnd=win32gui.FindWindowEx(None,None,None,arg[0])
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.5)
                except:
                    print("no",arg[0])
                    time.sleep(2)
                    pass
        else:
            print("w?")
        return win32gui.GetWindowRect(hwnd)[:2]
    def getPositionUS66(self,Purpose,WindowRec):
        # 获取控件的相对坐标
        win_patient=448,151
        Purpose=WindowRec[0]+Purpose[0]-win_patient[0],WindowRec[1]+Purpose[1]-win_patient[1]
        return Purpose
    def moveOnActive(self,box,windowname):
        # 移动到控件位置，并点击
        P,P1=self.getPositionUS66(box,self.getw1Window(windowname))
        self.MouseMoveAndClick(P,P1)
    def moveToInput(self,box,coment,windowname):
        # 移动到控件位置，并输入
        P,P1=self.getPositionUS66(box,self.getw1Window(windowname))
        self.MouseMoveTo(P,P1)
        self.Input(coment)
    def move(self,window_size,x,y):
        '''
        移动光标
        :param window_size: 屏幕分辨率,tuple
        :param x:
        :param y:
        :return:
        '''
        SW,SH=window_size
        mW=int(x*65535/SW)
        mH=int(y*65535/SH)
        win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE|win32con.MOUSEEVENTF_MOVE,mW,mH,0,0)
        time.sleep(0.2)
    def Ctrl_S(self,num):
        # 常用快捷键
        for x in range(0,num):
            self.PressKey("Ctrl","S")
            time.sleep(0.8)
    def Ctrl_E(self):
        # 常用快捷键
        self.PressKey("Ctrl","E")
    def Duuu(self):
        # 发出‘嘟’的一声
        windll.kernel32.Beep()

if __name__=="__main__":
    k=WisonicKey()
    k.getw1Window('钉钉')
    k.Duuu()