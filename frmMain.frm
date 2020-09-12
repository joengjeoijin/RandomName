VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "随机抽号"
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "控制台："
      BeginProperty Font 
         Name            =   "宋体"
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
         Caption         =   "下一个"
         BeginProperty Font 
            Name            =   "宋体"
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
      Caption         =   "抽选结果："
      BeginProperty Font 
         Name            =   "宋体"
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
            Name            =   "宋体"
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
Private Declare Function timeGetTime Lib "winmm.dll" () As Long '等待用
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
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置


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
    RndNum = Int(Rnda * Names + 1)
    lblName.Caption = AllNA(1, RndNum)
    Speaker.Speak "请" & AllNA(2, RndNum) & "同学答题。", 1
End Sub

Private Sub Form_Load()
    
    ReDim AllNA(2, 0)
    Me.Visible = True
    Me.Caption = "Loading"
    DoEvents
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Randomize
    Set Speaker = CreateObject("SAPI.SpVoice")
    Speaker.Volume = 100
    Speaker.Rate = -0.9
    
    
    Call ReadXLS
    
    Me.Caption = "随机抽号"
    Frame1.Visible = True
    ImgLoading.Visible = False
    Me.Picture = LoadPicture()
End Sub

Sub ReadMDB()
    Set DA = New Connection
    DA.CursorLocation = adUseClient
    DA.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\StudentName.MDB;User ID=admin; Password=; Jet OLEDB:Database Password=3.14"
    Set NA = New Recordset
    NA.Open "select * from CLASS", DA, adOpenStatic, adLockOptimistic
    NA.MoveFirst
    Do While Not NA.EOF
        ReDim Preserve AllNA(2, Names + 1)
        AllNA(1, UBound(AllNA, 2)) = NA("姓名")
        AllNA(2, UBound(AllNA, 2)) = NA("姓名")
        On Error Resume Next
        If NA("读音") <> "" Then
            AllNA(2, UBound(AllNA, 2)) = NA("读音")
        End If
        Names = Names + 1
        NA.MoveNext
    Loop
    On Error GoTo 0
End Sub

Sub ReadXLS()
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\StudentName.xlsx")
    xlApp.Visible = True
    Set xlSheet = xlBook.Worksheets("Names")
    Text = " "
    Do Until Text = ""
        ReDim Preserve AllNA(2, Names + 1)
        AllNA(1, UBound(AllNA, 2)) = xlApp.Cells(Names + 1, 1)
        AllNA(2, UBound(AllNA, 2)) = xlApp.Cells(Names + 1, 2)
        If AllNA(2, UBound(AllNA, 2)) = "" Then AllNA(2, UBound(AllNA, 2)) = AllNA(1, UBound(AllNA, 2))
        Text = AllNA(1, UBound(AllNA, 2))
        Debug.Print AllNA(1, UBound(AllNA, 2)), AllNA(2, UBound(AllNA, 2))
        Names = Names + 1
    Loop
    ReDim Preserve AllNA(2, UBound(AllNA, 2) - 1)
    Names = Names - 1
    xlBook.Close (True)
    xlApp.Quit
    Set xlApp = Nothing
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
End Sub

Public Sub lblName_Click()
    Call Command1_Click
End Sub
