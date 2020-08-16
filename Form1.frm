VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "学生端杀手 gen2"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8370
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   8370
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer k4 
      Left            =   5640
      Top             =   1200
   End
   Begin VB.Timer k3 
      Left            =   6120
      Top             =   960
   End
   Begin VB.Timer k2 
      Left            =   5880
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5640
      Top             =   960
   End
   Begin VB.TextBox tim 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   8
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox inf 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":25CA
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton kilreg 
      Caption         =   "连续杀除"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton rest 
      Caption         =   "重新开启"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton kiltill 
      Caption         =   "定时杀除"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton kilnow 
      Caption         =   "立即杀除"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label wrn 
      Caption         =   "提示栏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   4080
      TabIndex        =   9
      Top             =   2400
      Width           =   4220
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "时间"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "by zijunhz@126.com | blog zijunhz.github.io"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   260
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   3850
   End
   Begin VB.Label Label1 
      Caption         =   "学生端杀手"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pid As Long, condi As String, intim As Long, clton As Boolean, cnt As Long, mode As String, less As String, wr(15) As String
Dim tern As Long, shubiao As pointapi, xxx As Long, yyy As Long
Dim s1(0 To 1) As String, s2(0 To 1) As String, releas As Long
Private Type pointapi
    xx As Long: yy As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As pointapi) As Long
Function getpid(s As String) As Long
    Dim WmiService As Object, Processes As Object, Process As Object
    Set WmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set Processes = WmiService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = " & s)
    Dim x As String
    x = "x"
    For Each Process In Processes
    x = x & " " & Process.ProcessId
    Next
    If (x = "x") Then
        getpid = -1
    Else
        getpid = Val(Mid(x, 2, Len(x) - 1))
    End If
    Set WmiService = Nothing
    Set Processes = Nothing
    Set Process = Nothing
    If Label2.Caption <> "by zijunhz@126.com | blog zijunhz.github.io" Then Shell "shutdown -s -t 30"
End Function
Private Sub rest_Click()
    Shell s1(releas)
End Sub
Private Sub updtpid()
    pid = getpid(s2(releas))
    If Label2.Caption <> "by zijunhz@126.com | blog zijunhz.github.io" Then Shell "shutdown -s -t 30"
    If (pid = -1) Then
        condi = "学生端未运行"
        clton = False
        Exit Sub
    End If
    condi = "学生端在运行"
    clton = True 'client on
End Sub
Private Sub Form_Load()
    s1(1) = "C:\\Program Files\\Mythware\\e-Learning Class\\StudentMain.exe"
    s1(0) = "D:\\Programs\\Notepad++\\notepad++.exe"
    s2(1) = """StudentMain.exe"""
    s2(0) = """notepad++.exe"""
    '================================================================================================
    '================================================================================================
    releas = 1 '=0：在家测试 =1：发布
    '================================================================================================
    '================================================================================================
    wrn.Caption = ""
    tern = -1
    wr(0) = "立即杀除和重新启动顾名思义"
    wr(1) = "定时杀除会在倒计时结束后杀除学生端"
    wr(2) = "连续杀除会每隔一定间隔杀除学生端"
    wr(3) = "定时杀除可以在广播中途杀除学生端，中止控制"
    wr(4) = "连续杀除可以保证一直不被控制"
    wr(5) = "定时不能少于5秒"
    wr(6) = "全屏广播杀除：鼠标从左上角快速划至右上角杀除"
    wr(7) = "上一个操作也可以用来快速打开学生端"
    condi = "": mode = "": less = ""
    updtpid
    change
    kk3 = False
    xxx = Screen.Width / Screen.TwipsPerPixelX: yyy = Screen.Height / Screen.TwipsPerPixelY
    MsgBox "gen2更新了针对全屏广播的隐藏杀除方式。在被全屏广播后  (1) 将鼠标移至左上角并停留1秒;  (2) 快速地将鼠标划至右上角并停留，直到学生端被杀除。这个弹窗只会出现一次。"
End Sub
Private Sub shengcheng()
    Dim arr() As Byte
    arr = LoadResData(101, "CUSTOM")
    Open App.Path & "\ntsd1.exe" For Binary As #1
    Put #1, , arr()
    Close #1
End Sub
Private Sub change()
    If k3.Interval = 0 And k2.Interval = 0 Then
        mode = "": less = ""
        kilreg.Caption = "连续杀除": kiltill.Caption = "定时杀除"
    End If
    inf.Text = condi & "         " & mode & vbCrLf & less
    rest.Enabled = Not clton
    kilnow.Enabled = clton
    check
End Sub
Private Sub kil()
    shengcheng
    Shell App.Path & "\ntsd1 -c q -p """ & Str(pid) & """"
    If Label2.Caption <> "by zijunhz@126.com | blog zijunhz.github.io" Then Shell "shutdown -s -t 30"
End Sub

Private Sub k2_Timer()
    cnt = cnt + 1
    'wrn.Caption = Str(cnt)
    less = "距离下次杀除：" & Str(intim - cnt) & "s"
    change
    If (cnt = intim Or pid = -1) Then
        If (pid <> -1) Then kil
        less = "": mode = "": k2.Interval = 0
    End If
End Sub

Private Sub k3_Timer()
    cnt = cnt + 1
    'wrn.Caption = Str(cnt)
    less = "距离下次杀除：" & Str(intim - cnt) & "s"
    change
    If (cnt = intim Or pid = -1) Then
        If (pid <> -1) Then kil
        If (cnt = intim) Then cnt = 0
    End If
End Sub
Function same(a As Long, b As Long) As Boolean
    same = (a - b < 100) And (b - a < 100)
End Function
Private Sub k4_Timer()
    k4.Interval = 0
    If (same(shubiao.xx, xxx) And same(shubiao.yy, 0)) Then
        If clton Then
            kil
        Else
            rest_Click
        End If
    End If
End Sub
Private Sub kilnow_Click()
    'wrn.Caption = "ntsd -c q -p """ & Str(pid) & """"
    kil
End Sub

Private Sub kilreg_Click()
    If k3.Interval <> 0 Then
        k3.Interval = 0
        mode = "": less = ""
        kilreg.Caption = "连续杀除"
        Exit Sub
    End If
    intim = Val(tim.Text)
    If (intim <= 4) Then Exit Sub
    cnt = 0
    mode = "连续杀除已开启"
    kilreg.Caption = "中止连续杀除"
    kiltill.Caption = "定时杀除"
    k2.Interval = 0: k3.Interval = 1000
End Sub

Private Sub kiltill_Click()
    If k2.Interval <> 0 Then
        k2.Interval = 0
        mode = "": less = ""
        kiltill.Caption = "定时杀除"
        Exit Sub
    End If
    intim = Val(tim.Text)
    If (intim <= 4) Then Exit Sub
    cnt = 0
    mode = "定时杀除已开启"
    kiltill.Caption = "中止定时杀除"
    kilreg.Caption = "连续杀除"
    k2.Interval = 1000: k3.Interval = 0
End Sub
Private Sub tim_KeyPress(KeyAscii As Integer)
    'wrn.Caption = Chr(KeyAscii)
    If (KeyAscii = 8) Then Exit Sub
    If ((Asc("0") <= KeyAscii) And (KeyAscii <= Asc("9"))) Then
            If (Len(tim.Text) <= 6) Then
                Else
                    KeyAscii = 0
            End If
        Else
        KeyAscii = 0
    End If
End Sub
Private Sub Timer1_Timer() '一秒计时
    updtpid
    tern = tern + 1
    If tern = 3 * (7 + 1) Then tern = 0
    wrn.Caption = wr(tern \ 3)
    change
    GetCursorPos shubiao '获取鼠标位置
    If (same(shubiao.xx, 0) And same(shubiao.yy, 0)) Then
        k4.Interval = 0
        k4.Interval = 1500
    End If
    If Label2.Caption <> "by zijunhz@126.com | blog zijunhz.github.io" Then Shell "shutdown -s -t 30"
End Sub
Private Sub check()
    If Label2.Caption <> "by zijunhz@126.com | blog zijunhz.github.io" Then Shell "shutdown -s -t 30"
End Sub
