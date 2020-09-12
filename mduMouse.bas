Attribute VB_Name = "mduMouse"
Option Explicit

Public AllNA()
Public Names
Public Speaker
Public Declare Function timeGetTime Lib "winmm.dll" () As Long '等待用
Dim ScoredTime

'=============模块代码==================
Public Const WH_MOUSE = 7 '本地钩子
Public Const WH_MOUSE_LL = 14 '全局钩子
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const WM_LBUTTONDOWN = &H201 '窗口中按下鼠标左键
Public Const WM_LBUTTONUP = &H202 '窗口中松开鼠标左键
Public Const WM_MOUSEMOVE = &H200 '窗口中移动鼠标
Public Const WM_RBUTTONDOWN = &H204 '窗口中按下鼠标右键
Public Const WM_RBUTTONUP = &H205 '窗口中松开鼠标右键
Public Const WM_MOUSEWHEEL = &H20A '鼠标滚轮
Public Const WM_NCLBUTTONDOWN = &HA1 '窗口标题栏中按下鼠标左键
Public Const WM_NCLBUTTONUP = &HA2 '窗口标题栏中左开鼠标左键
Public Const WM_NCMOUSEMOVE = &HA0 '窗口标题栏中移动鼠标
Public Const WM_NCRBUTTONDOWN = &HA4 '窗口标题栏中按下鼠标右键
Public Const WM_NCRBUTTONUP = &HA5 '窗口标题栏中松开鼠标右键

Public hHook As Long
Public Type POINTAPI
X As Long
Y As Long
End Type

Type MSLLHOOKSTRUCT
pt As POINTAPI '相对于屏幕左上角的坐标x,y
mouseData As Long '鼠标数据
flags As Long '标记
time As Long '时间戳
dwExtraInfo As Long '其他信息
End Type

Type MOUSEHOOKSTRUCT
pt As POINTAPI '相对于屏幕左上角的坐标x,y
hwnd As Long '鼠标光标下窗口的句柄
wHitTestCode As Long '鼠标光标在窗口中的位置，标题栏、左边框、右边框，下边框。。。
dwExtraInfo As Long '其他信息，通常为0
End Type
Dim oMouseHookStruct As MSLLHOOKSTRUCT

Public Function MouseHookProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
CopyMemory oMouseHookStruct, ByVal lParam, Len(oMouseHookStruct)
'Debug.Print "当前鼠标位置-x:" & oMouseHookStruct.pt.X & "; y:" & oMouseHookStruct.pt.Y
Select Case wParam
Case WM_LBUTTONDOWN
'Debug.Print "左键按下"
    If timeGetTime < ScoredTime + 1000 Then GoTo Nex
    ScoredTime = timeGetTime
    If oMouseHookStruct.pt.X * Screen.TwipsPerPixelX - frmMain.Left > 0 And oMouseHookStruct.pt.Y * Screen.TwipsPerPixelY - frmMain.Top > 0 And oMouseHookStruct.pt.X * Screen.TwipsPerPixelX - frmMain.Left - frmMain.Width < 0 And oMouseHookStruct.pt.Y * Screen.TwipsPerPixelY - frmMain.Top - frmMain.Height < 0 Then
        frmMain.lblName.Caption = AllNA(Int(Rnd * Names + 1))
        Speaker.Speak "请" & frmMain.lblName.Caption & "同学答题。", 1
    End If
Case WM_NCLBUTTONDOWN
UnhookWindowsHookEx hHook
Unload frmMain
Case WM_LBUTTONUP, WM_NCLBUTTONUP
Debug.Print "左键弹起"
Case WM_RBUTTONDOWN, WM_NCRBUTTONDOWN
Debug.Print "右键按下"
Case WM_RBUTTONUP, WM_NCRBUTTONUP
Debug.Print "右键弹起"
Case WM_MOUSEMOVE, WM_NCMOUSEMOVE
Debug.Print "鼠标移动"
Case WM_MOUSEWHEEL
Debug.Print "鼠标滚轮"
End Select
Nex:
MouseHookProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
End Function
