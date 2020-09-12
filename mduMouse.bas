Attribute VB_Name = "mduMouse"
Option Explicit

Public AllNA()
Public Names
Public Speaker
Public Declare Function timeGetTime Lib "winmm.dll" () As Long '�ȴ���
Dim ScoredTime

'=============ģ�����==================
Public Const WH_MOUSE = 7 '���ع���
Public Const WH_MOUSE_LL = 14 'ȫ�ֹ���
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const WM_LBUTTONDOWN = &H201 '�����а���������
Public Const WM_LBUTTONUP = &H202 '�������ɿ�������
Public Const WM_MOUSEMOVE = &H200 '�������ƶ����
Public Const WM_RBUTTONDOWN = &H204 '�����а�������Ҽ�
Public Const WM_RBUTTONUP = &H205 '�������ɿ�����Ҽ�
Public Const WM_MOUSEWHEEL = &H20A '������
Public Const WM_NCLBUTTONDOWN = &HA1 '���ڱ������а���������
Public Const WM_NCLBUTTONUP = &HA2 '���ڱ���������������
Public Const WM_NCMOUSEMOVE = &HA0 '���ڱ��������ƶ����
Public Const WM_NCRBUTTONDOWN = &HA4 '���ڱ������а�������Ҽ�
Public Const WM_NCRBUTTONUP = &HA5 '���ڱ��������ɿ�����Ҽ�

Public hHook As Long
Public Type POINTAPI
X As Long
Y As Long
End Type

Type MSLLHOOKSTRUCT
pt As POINTAPI '�������Ļ���Ͻǵ�����x,y
mouseData As Long '�������
flags As Long '���
time As Long 'ʱ���
dwExtraInfo As Long '������Ϣ
End Type

Type MOUSEHOOKSTRUCT
pt As POINTAPI '�������Ļ���Ͻǵ�����x,y
hwnd As Long '������´��ڵľ��
wHitTestCode As Long '������ڴ����е�λ�ã�����������߿��ұ߿��±߿򡣡���
dwExtraInfo As Long '������Ϣ��ͨ��Ϊ0
End Type
Dim oMouseHookStruct As MSLLHOOKSTRUCT

Public Function MouseHookProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
CopyMemory oMouseHookStruct, ByVal lParam, Len(oMouseHookStruct)
'Debug.Print "��ǰ���λ��-x:" & oMouseHookStruct.pt.X & "; y:" & oMouseHookStruct.pt.Y
Select Case wParam
Case WM_LBUTTONDOWN
'Debug.Print "�������"
    If timeGetTime < ScoredTime + 1000 Then GoTo Nex
    ScoredTime = timeGetTime
    If oMouseHookStruct.pt.X * Screen.TwipsPerPixelX - frmMain.Left > 0 And oMouseHookStruct.pt.Y * Screen.TwipsPerPixelY - frmMain.Top > 0 And oMouseHookStruct.pt.X * Screen.TwipsPerPixelX - frmMain.Left - frmMain.Width < 0 And oMouseHookStruct.pt.Y * Screen.TwipsPerPixelY - frmMain.Top - frmMain.Height < 0 Then
        frmMain.lblName.Caption = AllNA(Int(Rnd * Names + 1))
        Speaker.Speak "��" & frmMain.lblName.Caption & "ͬѧ���⡣", 1
    End If
Case WM_NCLBUTTONDOWN
UnhookWindowsHookEx hHook
Unload frmMain
Case WM_LBUTTONUP, WM_NCLBUTTONUP
Debug.Print "�������"
Case WM_RBUTTONDOWN, WM_NCRBUTTONDOWN
Debug.Print "�Ҽ�����"
Case WM_RBUTTONUP, WM_NCRBUTTONUP
Debug.Print "�Ҽ�����"
Case WM_MOUSEMOVE, WM_NCMOUSEMOVE
Debug.Print "����ƶ�"
Case WM_MOUSEWHEEL
Debug.Print "������"
End Select
Nex:
MouseHookProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
End Function
