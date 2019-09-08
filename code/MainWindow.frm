VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "嘤悦 II"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13350
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   890
   StartUpPosition =   2  '屏幕中心
   Begin MSWinsockLib.Winsock MainSock 
      Left            =   300
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer FPSPrinter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12750
      Top             =   150
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyGame As New GameManager, MyFont As Fonts
Dim FPS As Long, LastFPS As Long
Dim DebugDraw As Images, DebugTemp As Images
Dim LastMouseTime As Currency, LeaveDialog As Boolean
Dim MouseVisible As Boolean
Dim DrawingUse As Currency
Dim ClosePower As Boolean
Dim LastKey As Integer
Sub DrawGame()
    
    If mNowShow <> "EditPage" And mNowShow <> "Notify" And mNowShow <> "PlayPage" Then
        If GameBGM.PlayState <> Playing Then
            If UBound(SongList) > 0 Then
                PlayRandomSong
            End If
        End If
    End If
    If mNowShow <> "EditPage" Then
        Dim TargetRate As Single
        TargetRate = IIf(ModPower(1), 1.5, 1)
        If GameBGM.Rate <> TargetRate Then GameBGM.SetPlayRate TargetRate
    End If
    
    MyGame.Display

    FPS = FPS + 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If LastKey = KeyCode Then Exit Sub
    LastKey = KeyCode
    If mNowShow = "EditPage" Then EditPage.KeyDown KeyCode
    If mNowShow = "PlayPage" Then PlayPage.KeyDown KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If NowEdit <> -1 Then
        If KeyAscii = vbKeyBack Then
            If Len(EditBoxText(NowEdit)) > 0 Then
                EditBoxText(NowEdit) = Left(EditBoxText(NowEdit), Len(EditBoxText(NowEdit)) - 1)
            Else
                VBA.Beep
            End If
        ElseIf KeyAscii <> vbKeyReturn Then
            EditBoxText(NowEdit) = EditBoxText(NowEdit) & Chr(KeyAscii)
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    LastKey = 0
    If mNowShow = "EditPage" Then EditPage.KeyUp KeyCode
    If mNowShow = "PlayPage" Then PlayPage.KeyUp KeyCode
    PressLock = -1
End Sub

Private Sub Form_Load()

    PressLock = -1

    MouseVisible = True
    PlayIcelolly Me, App.Path, True
    ReDim SongList(0)
    LockX = -1
    
    For i = 0 To 3
        Set ObjImg(i) = PlayImg.Image("object" & i & ".png")
    Next
    For i = 0 To 3
        Set ObjImg2(i) = PlayImg.Image("object2" & i & ".png")
    Next
    Set ObjImg(4) = PlayImg.Image("now.png")
    
    Set EffectTipImg(0) = PlayImg.Image("excellent.png")
    Set EffectTipImg(1) = PlayImg.Image("great.png")
    Set EffectTipImg(2) = PlayImg.Image("okay.png")
    Set EffectTipImg(3) = PlayImg.Image("miss.png")
    
    If DebugMode = False Then MWndProc = SetWindowLongA(Me.hwnd, GWL_WNDPROC, AddressOf FunWndProc)
    
    Set MyFont = New Fonts
    MyFont.Create "微软雅黑"
    MyFont.SetFont
    
    Set MainPage = New MainPage
    Set EditPage = New EditPage
    Set Notify = New Notify
    Set ByePage = New ByePage
    Set SoloPage = New SoloPage
    Set PlayPage = New PlayPage
    Set DownPage = New DownPage
    Set Tasks = New Tasks
    Set SetPage = New SetPage
    
    MyGame.AddScreen MainPage
    MyGame.AddScreen EditPage
    MyGame.AddScreen ByePage
    MyGame.AddScreen SoloPage
    MyGame.AddScreen PlayPage
    MyGame.AddScreen DownPage
    MyGame.AddScreen SetPage
    MyGame.AddScreen Notify
    MyGame.AddScreen Tasks
    
    NowShow = "MainPage"
    
    Set DebugDraw = New Images: Set DebugTemp = New Images
    DebugDraw.Create Me.hdc, GWW, GWH
    DebugTemp.Create Me.hdc, GWW, GWH
    
    SetMainBack App.Path & "\assets\background\white.png"
    
    GetSongList
    If UBound(SongList) = 0 Then
        LoadMapFile App.Path & "\assets\example.mumap"
        GetSongList
    End If
    If UBound(SongList) > 0 Then
        PlayRandomSong
    End If
    
    FPSPrinter.Enabled = True
    
    Set Sock = MainSock
    
    If DebugMode = False Then
        Dim FuckSuccess As Boolean
        FuckSuccess = True
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD) <> 0)
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(WM_COPYGLOBALDATA, MSGFLT_ADD) <> 0)
        FuckSuccess = FuckSuccess And (ChangeWindowMessageFilter(404233, MSGFLT_ADD) <> 0)
        DragAcceptFiles Me.hwnd, 1
    End If
    
    EditBoxText(1) = "userid"
    EditBoxText(2) = "password"
    
    GameOpenTime = GetTickCount
    
    SaveSetting "MuingII", "RunTime", "Window", Me.hwnd
    
    Dim FileReg As Object
    Set FileReg = CreateObject("Wscript.Shell")
    
    On Error Resume Next
    Err.Clear
    
    Dim tempReg As String
    tempReg = FileReg.RegRead("HKCR\MuingII.Map\")
    
    If Err.Number <> 0 Then
        FileReg.RegWrite "HKCR\MuingII.Map\", "嘤悦2 谱子包", "REG_SZ"
        FileReg.RegWrite "HKCR\MuingII.Map\DefaultIcon\", """" & App.Path & "\assets\icon\map.ico" & """", "REG_SZ"
        FileReg.RegWrite "HKCR\MuingII.Map\shell\open\command\", """" & App.Path & "\" & App.EXEName & ".exe" & """ ""%1""", "REG_SZ"
        FileReg.RegWrite "HKCR\.mumap\", "MuingII.Map", "REG_SZ"
        
        FileReg.RegWrite "HKCR\MuingII.Level\", "嘤悦2 谱子信息", "REG_SZ"
        FileReg.RegWrite "HKCR\MuingII.Level\DefaultIcon\", """" & App.Path & "\assets\icon\level.ico" & """", "REG_SZ"
        'FileReg.RegWrite "HKCR\MuingII.Level\shell\open\command\", "", "REG_SZ"
        FileReg.RegWrite "HKCR\.mu\", "MuingII.Level", "REG_SZ"
    End If
    
    Err.Clear
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UpdateClickTest x, y, 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MouseState <> 2 Then UpdateClickTest x, y, IIf(Button = 0, 0, 1)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UpdateClickTest x, y, 2
End Sub

Public Sub FileDrop(List As String)
    Dim data() As String
    data = Split(List, vbCrLf)
    
    If mNowShow = "EditPage" Then
        EditPage.FileDrop data(0)
    End If
    If mNowShow = "MainPage" Or mNowShow = "SoloPage" Then
        MainPage.FileDrop List
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "MuingII", "RunTime", "Window", 0
    If ClosePower = False Then
        If UserID <> 0 And Connected And Newest Then
            Send "adduptime*" & Int((GetTickCount - GameOpenTime) / 1000)
        End If
        ClosePower = True: MenuSnd2.Play
        GameCore.FadePage "ByePage"
        Cancel = 1
    Else
        CloseTime = True
    End If
End Sub

Private Sub FPSPrinter_Timer()
    Me.Caption = "嘤悦 II"

    For i = 0 To UBound(HistoryFPS) - 1
        HistoryFPS(i) = HistoryFPS(i + 1)
    Next
    HistoryFPS(UBound(HistoryFPS)) = FPS
    
    For i = 0 To UBound(HistoryDraw) - 1
        HistoryDraw(i) = HistoryDraw(i + 1)
    Next
    HistoryDraw(UBound(HistoryDraw)) = LastDrawC
    LastDrawC = 0
    
    LastFPS = FPS
    
    FPS = 0
End Sub

Private Sub MainSock_Close()
    GameNotify.Message MuingServer, "服务器断开了连接。", "噢"
    GiveupConnect = True
    Connected = False
    UserID = 0
End Sub

Private Sub MainSock_Connect()
    Send "check*2019010501"
End Sub

Private Sub MainSock_DataArrival(ByVal bytesTotal As Long)
    If DebugMode = False Then On Error Resume Next

    Dim Command As String
    Dim Cmd() As String, Cmd2() As String
    
    MainSock.GetData Command
    Cmd = Split(Command, Chr(404))
    
    For i = 0 To UBound(Cmd) - 1
        Cmd2 = Split(Cmd(i), "*")
        
        If Cmd2(0) = "notify" Then
            GameNotify.Message MuingServer, "来自服务器的消息：" & vbCrLf & Cmd2(1), "确定"
        End If
        
        If Cmd2(0) = "t" Then
            CLate = (GetTickCount - Val(Cmd2(1)))
        End If
        
        If Cmd2(0) = "check" Then
            Connected = True
            If Cmd2(1) = "yes" Then
                Newest = True
                If GameSave.RSave("UserID") <> "" Then
                    Send "log*" & GameSave.RSave("UserID") & "*" & GameSave.RSave("Password")
                    MsgCheck = ""
                    Do While MsgCheck = ""
                        DoEvents
                    Loop
                    Dim temp() As String
                    temp = Split(MsgCheck, "*")
                    If temp(1) <> "no" Then
                        UserName = temp(1): UserID = Int(Val(GameSave.RSave("UserID")))
                        LoginTime = temp(2): BaseTime = Val(temp(3))
                        Exp = Val(temp(4))
                    End If
                End If
            Else
                UpdateUrl = Cmd2(2): UpdateVersion = Cmd2(3)
                GameNotify.Message MuingUpdate, "你的游戏版本不是最新的，多人游戏被禁用，且你的分数不会被上传。", "噢"
                MainSock.close
            End If
        End If
        
        If Cmd2(0) = "checkmap" Then
            MapCheck = Cmd(i)
        End If
        
        If Cmd2(0) = "receivemap" Then
            MapList = Cmd(i)
        End If
        
        If Cmd2(0) = "log" Or Cmd2(0) = "reg" Or Cmd2(0) = "rank" Then MsgCheck = Cmd(i)
    Next
End Sub

Private Sub MainSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    GameNotify.Message MuingNoConnected, "连接到服务器时出错。" & vbCrLf & Description & "(" & Number & ")", "噢"
    GiveupConnect = True
    Connected = False
    UserID = 0
End Sub


