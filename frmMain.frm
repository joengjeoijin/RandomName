VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2745
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   1410
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "����̨��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton Command1 
         Caption         =   "��һ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ѡ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   795
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2280
      End
   End
   Begin VB.Image ImgLoading 
      Appearance      =   0  'Flat
      Height          =   1800
      Left            =   0
      Picture         =   "frmMain.frx":1245
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2835
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DA
Dim NA
Dim LoadingNum
Dim AllNA()
Dim Names
Dim Speaker
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '�ȴ���
Dim ScoredTime


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H90000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Dim hwng


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��


Private Sub Command1_Click()
    Speaker.pause
    Set Speaker = CreateObject("SAPI.SpVoice")
    Speaker.Volume = 100
    Speaker.Rate = -0.9
    Dim Rnda As Double
    Rnda = timeGetTime / 1000
    Rnda = Rnda - Int(Rnda)
    Rnda = Rnd + Rnda
    Rnda = Rnda - Int(Rnda)
    lblName.Caption = AllNA(Int(Rnda * Names + 1))
    Speaker.Speak "��" & lblName.Caption & "ͬѧ���⡣", 1
End Sub

Private Sub Form_Load()
    
    ReDim AllNA(0)
    Me.Visible = True
    Me.Caption = "Loading"
    DoEvents
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Randomize
    'hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHookProc, App.hInstance, 0)
    Set Speaker = CreateObject("SAPI.SpVoice")
    Speaker.Volume = 100
    Speaker.Rate = -0.9
    Set DA = New Connection
    DA.CursorLocation = adUseClient
    DA.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\StudentName.MDB;User ID=admin; Password=; Jet OLEDB:Database Password=3.14"
    Set NA = New Recordset
    NA.Open "select * from CLASS", DA, adOpenStatic, adLockOptimistic
    NA.MoveFirst
    Do While Not NA.EOF
        ReDim Preserve AllNA(Names + 1)
        AllNA(UBound(AllNA, 1)) = NA("����")
        Names = Names + 1
        NA.MoveNext
    Loop
    Me.Caption = "������"
    Frame1.Visible = True
    ImgLoading.Visible = False
    Me.Picture = LoadPicture()
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
    ImgLoading.Picture = LoadPicture(App.Path & "\Rotate" & CStr(LoadingNum Mod 7 + 1) & ".ico")
    LoadingNum = LoadingNum + 1
End Sub

Public Sub Form_Unload(Cancel As Integer)
    'UnhookWindowsHookEx hHook
End Sub

Public Sub lblName_Click()
    Call Command1_Click
End Sub
