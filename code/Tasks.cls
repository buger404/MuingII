VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images, Back As New Images
Dim PageList(4) As String
Dim AlphaId As Integer
Dim Remember As Boolean
Dim LastPage As String
Public LoginMode As Boolean, LoginM As Integer, ReNameMode As Boolean, MusicMode As Boolean
Public UpdateMode As Boolean
Public Property Get Visible() As Boolean
    Visible = ShowTask
End Property
Sub DrawUpdateUI()
    MainImg.Image("updateframe.png").Present Draw.CompatibleDC, 250, 173
    
    GameFont.DrawText Draw.Graphics, 250, 173 + 27, 410, 40, "更新可用", argb(255, 255, 255, 255), far, 24, Bold
    GameFont.DrawText Draw.Graphics, 250, 173 + 27 + 35, 410, 40, UpdateVersion, argb(200, 255, 255, 255), far, 20, Regular
    
    GameFont.DrawText Draw.Graphics, 268, 173 + 125, 410, 40, "你的游戏不是最新版，所有联网功能将被禁用。", argb(255, 44, 41, 30), center, 16, Regular
    
    CtrlImg.ImageByIndex(1).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 538, 173 + 190, "立刻更新", argb(255, 255, 255, 255), 16, Regular, Button
    If IsClick Then ShellExecuteA MainWindow.hwnd, "open", UpdateUrl, "", "", SW_SHOW
End Sub
Sub DrawMusicUI()
    MainImg.Image("songframe.png").Present Draw.CompatibleDC, 163, 152
    SongList(NowSong).SongPic.Present Draw.CompatibleDC, 170, 154
    
    With MainImg.Image("leftturn.png")
        .SetClickArea 185, 177
        .Present Draw.CompatibleDC, 185, 177, IIf(IsMouseIn, 255, 60)
        If IsClick And UBound(SongList) > 0 Then
            MenuSnd.Play
            NowSong = NowSong - 1
            If NowSong <= 0 Then NowSong = UBound(SongList)
            GameBGM.LoadMusic SongPath & "\" & SongList(NowSong).Path & "\music.mp3"
            GameBGM.Play
            SetMainBack SongPath & "\" & SongList(NowSong).Path & "\background.png"
            SoloPage.ChangeIndex = NowSong: SoloPage.ChangeTime = GetTick
            SoloPage.ScrollTo NowSong
            GameCore.FadePage mNowShow
        End If
    End With
    
    With MainImg.Image("rightturn.png")
        .SetClickArea 672, 177
        .Present Draw.CompatibleDC, 672, 177, IIf(IsMouseIn, 255, 60)
        If IsClick And UBound(SongList) > 0 Then
            MenuSnd.Play
            NowSong = NowSong + 1
            If NowSong > UBound(SongList) Then NowSong = 1
            GameBGM.LoadMusic SongPath & "\" & SongList(NowSong).Path & "\music.mp3"
            GameBGM.Play
            SetMainBack SongPath & "\" & SongList(NowSong).Path & "\background.png"
            SoloPage.ChangeIndex = NowSong: SoloPage.ChangeTime = GetTick
            SoloPage.ScrollTo NowSong
            GameCore.FadePage mNowShow
        End If
    End With
    
    GameFont.DrawText Draw.Graphics, 190, 274, 600, 30, "Playing " & SongList(NowSong).Info.Title & " ~", argb(255, 64, 64, 64), near, 16
    
    Dim Pos As Single
    Pos = GameBGM.Pos
    CtrlImg.ImageByIndex(5).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 210, 314, Pos, 0, 1, Regular, HScrollBar, , GameBGM.max
    If IsMouseIn And MouseState > 0 Then GameBGM.Pos = Pos
    
    GameFont.DrawText Draw.Graphics, 190, 349, 525, 30, Int(Pos / 60) & ":" & format(Pos Mod 60, "00") & " / " & Int(GameBGM.max / 60) & ":" & format(GameBGM.max Mod 60, "00"), argb(255, 64, 64, 64), far, 16
End Sub
Sub DrawLoginUI()
    OnlineImg.Image("loginframe.png").Present Draw.CompatibleDC, 5, 50
    
    DrawWave 12, 112, 353, 150, 70
    
    If UserID <> 0 Then
        If ReNameMode = False Then GameFont.DrawText Draw.Graphics, 35, 104, 280, 30, UserName, argb(255, 27, 27, 27), near, 16, Bold
        SetClickArea2 35, 104, 280, 30
        Dim MouseUp As Boolean
        MouseUp = IsMouseUp
        If ReNameMode = True And GetAsyncKeyState(VK_RETURN) < 0 Then MouseUp = True
        If MouseUp Then
            ReNameMode = Not ReNameMode
            If ReNameMode = False Then
                If UserName <> EditBoxText(3) And EditBoxText(3) <> "" Then
                    UserName = EditBoxText(3)
                    Send "rename*" & UserName
                End If
            Else
                EditBoxText(3) = UserName
            End If
        End If
        
        If ReNameMode Then CtrlImg.ImageByIndex(3).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 35, 104, 3, argb(255, 27, 27, 27), 16, Bold, EditBox
        
        GameFont.DrawText Draw.Graphics, 35, 134, 280, 30, "登陆时间    " & LoginTime, argb(180, 27, 27, 27), near, 14
        GameFont.DrawText Draw.Graphics, 35, 154, 280, 30, "游戏时间    " & Int(((GetTickCount - GameOpenTime) / 1000 + BaseTime) / 60) & ":" & format(Int((GetTickCount - GameOpenTime) / 1000) Mod 60, "00"), argb(180, 27, 27, 27), near, 14
        
        GameFont.DrawText Draw.Graphics, 35, 184, 280, 30, "Exp " & Val(Exp), argb(180, 27, 27, 27), near, 14
        
        CtrlImg.ImageByIndex(BMCtrl.BMNormalButton).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 245, 238, "登出", argb(255, 255, 255, 255), 14, Regular, Button
        If IsClick Then
            UserID = 0: UserName = "": Send "logout"
        End If
        Exit Sub
    End If
    
    If LoginM = 0 Then
        OnlineImg.Image("login.png").Present Draw.CompatibleDC, 255, 70
    Else
        OnlineImg.Image("join.png").Present Draw.CompatibleDC, 233, 70
    End If
    If LoginM = 0 Then
        CtrlImg.ImageByIndex(3).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 35, 124, 1, argb(255, 64, 64, 64), 16, Regular, EditBox
        If IsClick And EditBoxText(1) = "userid" Then EditBoxText(1) = 0
    End If
    If EditBoxText(2) = "password" Then
        CtrlImg.ImageByIndex(3).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 35, IIf(LoginM = 0, 163, 124), 2, argb(255, 64, 64, 64), 16, Regular, EditBox
        If IsClick And EditBoxText(2) = "password" Then EditBoxText(2) = ""
    Else
        CtrlImg.ImageByIndex(3).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 35, IIf(LoginM = 0, 163, 124), 2, argb(255, 64, 64, 64), 16, Regular, EditBox, , 1
    End If

    If EditBoxText(1) <> "userid" Then
        EditBoxText(1) = Int(Val(EditBoxText(1)))
        If Len(EditBoxText(1)) > 5 Then EditBoxText(1) = Left(EditBoxText(1), 5)
    End If

    If LoginM = 0 Then
        CtrlImg.ImageByIndex(4).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 35, 210, "保持登录状态", argb(200, 27, 27, 27), 14, Regular, CheckBox, , Remember
    End If

    GameFont.DrawText Draw.Graphics, 285, 208, 100, 30, IIf(LoginM = 0, "注册一个", "返回登陆"), argb(255, 27, 27, 27), near, 14
    SetClickArea2 285, 208, 100, 30
    If IsClick Then
        If LoginM = 0 Then
            LoginM = 1: EditBoxText(2) = "password"
        Else
            LoginM = 0
        End If
    End If

    CtrlImg.ImageByIndex(BMCtrl.BMNormalButton).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 245, 238, IIf(LoginM = 0, "登陆", "注册"), argb(255, 255, 255, 255), 14, Regular, Button
    If IsClick Then
        If EditBoxText(2) = "" Then GameNotify.Message MuingServer, "密码不能为空。", "噢": Exit Sub
        If LoginM = 0 Then
            Send "log*" & Int(Val(EditBoxText(1))) & "*" & BMEA(EditBoxText(2))
        Else
            Send "reg*" & BMEA(EditBoxText(2))
        End If
        Sleep 1000
        MsgCheck = ""
        Do While MsgCheck = ""
            DoEvents
        Loop
        Dim temp() As String
        If LoginM = 0 Then
            temp = Split(MsgCheck, "*")
            If temp(1) = "no" Then
                GameNotify.Message MuingServer, "用户名或密码不正确！", "噫？"
            Else
                If Remember Then GameSave.WSave "UserID", Int(Val(EditBoxText(1))): GameSave.WSave "Password", BMEA(EditBoxText(2))
                LoginMode = False: MusicMode = False: mNowShow = LastPage
                GameNotify.Message MuingSuccess, "登陆成功", "好"
                UserName = temp(1): UserID = Int(Val(EditBoxText(1)))
                LoginTime = temp(2): BaseTime = Val(temp(3))
            End If
        Else
            temp = Split(MsgCheck, "*")
            LoginMode = False: MusicMode = False
            GameNotify.Message MuingSuccess, "注册成功，请牢记您的账号ID：" & vbCrLf & temp(1), "OK"
        End If
    End If
    
End Sub
Sub DrawWave(ByVal DX As Long, ByVal DY As Long, ByVal W As Long, ByVal H As Long, ByVal Deep As Long)
    Dim Width2 As Long, Height2 As Long, Path As Long, Brush As Long, Ox As Long
    Dim y As Long, x As Long
    Width2 = W: Height2 = H:  Ox = GetTickCount / 5
    
    GdipCreateLineBrush NewPointF(0, DY), NewPointF(0, DY + H), argb(255, 21, 234, 183), argb(40, 21, 234, 183), WrapModeTile, Brush
    GdipCreatePath FillModeWinding, Path
    For i = 0 To W Step 20
        x = i + Ox
        GdipAddPathLine Path, DX + x - Ox, DY + Sin(x / 150) * 30 + Height2 - Deep - 50, DX + x - Ox, DY + Sin(x / 150) * 30 + Height2 - Deep - 50
    Next
    x = W + Ox
    GdipAddPathLine Path, DX + x - Ox, DY + Sin(x / 150) * 30 + Height2 - Deep - 50, DX + x - Ox, DY + Sin(x / 150) * 30 + Height2 - Deep - 50
    
    GdipAddPathLine Path, DX + Width2, DY + Height2 - 1, DX, DY + Height2 - 1

    GdipFillPath Draw.Graphics, Brush, Path

    GdipDeletePath Path
    GdipDeleteBrush Brush
End Sub
Sub ChangePage(ByVal i As Integer)
    LoginMode = False: MusicMode = False

    If mNowShow = "EditPage" And EditPage.Page = 1 Then  '特殊情况
        If EditPage.Saved = False Then
            If GameNotify.Asks(MuingAsk, "您的谱子还未保存，继续退出吗？", "不要", "退出") = 0 Then Exit Sub
        End If
        EditPage.Page = 0
    End If
    
    If i = 1 Or i = 3 Then '需要连接网络的页面
        If Connected Then
            If Newest Then
                MenuSnd2.Play
                If i = 1 Then
                    GameNotify.Message MuingInformation, "It's not ready for use yet !", "..."
                Else
                    GameCore.FadePage PageList(i)
                    If i = 3 Then Send "getmap"
                End If
            Else
                GameNotify.Message MuingUpdate, "你的游戏版本不是最新的。", "..."
            End If
        ElseIf GiveupConnect = False Then
            MenuSnd2.Play: GameNotify.Message MuingServer, "请您等一等...服务器还没有连接上...", "..."
        Else
            MenuSnd2.Play: GameNotify.Message MuingNoConnected, "无法与服务器建立联系。", "..."
        End If
    Else
        MenuSnd2.Play: GameCore.FadePage PageList(i)
    End If
End Sub

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "Tasks"
    
    GdipGraphicsClear Draw.Graphics, 0
    If mNowShow = "LoginPage" Then Back.Present Draw.CompatibleDC, 0, 0: BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, HotAlpha
    If mNowShow = "MusicControl" Then Back.Present Draw.CompatibleDC, 0, 0: BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, HotAlpha
    If mNowShow = "UpdatePage" Then Back.Present Draw.CompatibleDC, 0, 0: BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, HotAlpha
    
    If mNowShow = "PlayPage" Or mNowShow = "Notify" Or mNowShow = "ByePage" Or (ShowTask = False) Then GoTo last
    
    '页面标签绘制
    Dim x As Long
    Dim Alpha As Long
    
    With PageImg.ImageByIndex(1)
        .SetClickArea -7, -3
        Alpha = IIf(IsMouseIn, 255, 100)
        If LoginMode Or MusicMode Then Alpha = 255
        LockMousePage = IIf(Alpha = 255, "Tasks", "")
        If Ani.GetTarget(AlphaId) <> Alpha Then Ani.ChangeTaskFast AlphaId, Alpha
        Alpha = Ani.GetValue(AlphaId)
        .Present Draw.CompatibleDC, -7, -3
    End With
    x = 82
    For i = 0 To UBound(PageList)
        If mNowShow = PageList(i) Then
            With PageImg.ImageByIndex(3 + i * 2)
                .Present Draw.CompatibleDC, x, 3
                PageImg.ImageByIndex(2).Present Draw.CompatibleDC, x + .Width / 2 - PageImg.ImageByIndex(2).Width / 2, 47
                x = x + PageImg.ImageByIndex(3 + i * 2).Width + 8
            End With
        Else
            With PageImg.ImageByIndex(3 + i * 2 + 1)
                .Present Draw.CompatibleDC, x - 3, 3
                .SetClickArea x, 2
                If IsClick Then '切换页面处理
                    ChangePage i
                End If
                x = x + PageImg.ImageByIndex(3 + i * 2).Width + 8
            End With
        End If
        
    Next
        
    '返回按钮
    If mNowShow <> "MainPage" Then
        With PageImg.ImageByIndex(3 + (UBound(PageList) + 1) * 2 + 2)
            .Present Draw.CompatibleDC, GWW - 39, 22
            .SetClickArea GWW - 39, 22
            If IsClick Then
                LoginMode = False: MusicMode = False: UpdateMode = False
                MenuSnd2.Play: GameCore.ChangePage "MainPage"
                If mNowShow = "EditPage" And EditPage.Page = 1 Then
                    If EditPage.Saved = False Then
                        If GameNotify.Asks(MuingAsk, "您的谱子还未保存，继续退出吗？", "不要", "退出") = 0 Then Exit Sub
                    End If
                    EditPage.Page = 0
                    GameCore.FadePage "EditPage"
                End If
            End If
            If IsMouseIn Then
                GameNotify.TipText = "返回：返回上一级页面。"
            End If
        End With
    End If
        
    '音乐按钮
    If mNowShow <> "EditPage" And EditPage.Page <> 1 Then
        If MusicMode Then BackImg.ImageByIndex(3).PresentWithClip Draw.CompatibleDC, GWW - 100, 0, 0, 0, 50, 55, 80
        With PageImg.ImageByIndex(3 + (UBound(PageList) + 1) * 2 + 1)
            .Present Draw.CompatibleDC, GWW - 82, 22
            .SetClickArea GWW - 82, 22
            If IsClick Then
                LoginMode = False: UpdateMode = False
                MusicMode = Not MusicMode
                MenuSnd.Play
                If MusicMode Then BlurTo Back.CompatibleDC, 20
                If mNowShow <> "MusicControl" And mNowShow <> "LoginPage" And mNowShow <> "UpdatePage" Then LastPage = mNowShow
                GameCore.FadePage IIf(MusicMode, "MusicControl", LastPage)
            End If
            If IsMouseIn Then
                GameNotify.TipText = "播放器：控制音乐的进度~"
            End If
        End With
    End If
    
    '更新按钮
    If UpdateUrl <> "" Then
        If UpdateMode Then BackImg.ImageByIndex(3).PresentWithClip Draw.CompatibleDC, GWW - 100 - 44, 0, 0, 0, 50, 55, 80
        With PageImg.ImageByIndex(3 + (UBound(PageList) + 1) * 2)
            .Present Draw.CompatibleDC, GWW - 82 - 43, 22
            .SetClickArea GWW - 82 - 43, 22
            If IsClick Then
                LoginMode = False: MusicMode = False
                UpdateMode = Not UpdateMode
                MenuSnd.Play
                If UpdateMode Then BlurTo Back.CompatibleDC, 20
                If mNowShow <> "MusicControl" And mNowShow <> "LoginPage" And mNowShow <> "UpdatePage" Then LastPage = mNowShow
                GameCore.FadePage IIf(UpdateMode, "UpdatePage", LastPage)
            End If
            If IsMouseIn Then
                GameNotify.TipText = "更新：更新到最新版本以使用联机功能。"
            End If
        End With
    End If
    
    If LoginMode Then
        DrawLoginUI
    End If
    
    If MusicMode Then
        DrawMusicUI
    End If
    
    If UpdateMode Then
        DrawUpdateUI
    End If
    
    If LoginMode Then BackImg.ImageByIndex(3).PresentWithClip Draw.CompatibleDC, 0, 0, 0, 0, 87, 55, 80
    If UserID = 0 Then
        With MainImg.Image("nobody.png")
            .Present Draw.CompatibleDC, 12, 20
            .SetClickArea 30, 30
        End With
    Else
        With MainImg.Image("rename.png")
            .Present Draw.CompatibleDC, 12, 20
            .SetClickArea 30, 30
        End With
    End If
    
    If IsClick Then
        If Connected Then
            If Newest Then
                MenuSnd2.Play: LoginMode = Not LoginMode
                MusicMode = False: UpdateMode = False
                If LoginMode = True Then BlurTo Back.CompatibleDC, 20
                If mNowShow <> "MusicControl" And mNowShow <> "LoginPage" And mNowShow <> "UpdatePage" Then LastPage = mNowShow
                GameCore.FadePage IIf(LoginMode, "LoginPage", LastPage)
            Else
                GameNotify.Message MuingUpdate, "你的游戏不是最新的。", "..."
            End If
        ElseIf GiveupConnect = False Then
            MenuSnd2.Play: GameNotify.Message MuingServer, "请您等一等...服务器还没有连接上...", "..."
        Else
            MenuSnd2.Play: GameNotify.Message MuingNoConnected, "无法与服务器建立联系。", "..."
        End If
    End If
    
last:
    
    If mNowShow <> "MainPage" And mNowShow <> "Notify" And mNowShow <> "PlayPage" And Settings(MuingII_Settings.MuingShowHot) = 0 Then
        Dim MusicD() As Single
        MusicD = GameBGM.GetMusicBar
        
        Dim Hot As Long, HotW As Long, NowHot As Long
        HotW = Int(GWW / 30)
        
        If GameBGM.PlayState = Playing Then
            For i = 0 To UBound(MusicD)
                If i <= 30 Then
                    Hot = MusicD(i) / 8
                    BackImg.ImageByIndex(1).PresentWithClip Draw.CompatibleDC, i * HotW, GWH - Hot, 0, 0, HotW, Hot, (70 + Hot / (GWH) * 180) * IIf(Alpha = 255, 1, 0.5)
                End If
                NowHot = NowHot + MusicD(i)
            Next
        End If
        HotAlpha = 110 - (NowHot / 1000) - EffectLig
        
        Draw.Present GameDraw.CompatibleDC, 0, 0, IIf(Alpha = 0, 100, Alpha)
    End If

    If Settings(MuingII_Settings.MuingShowHot2) = 1 Then HotAlpha = 0

    If Settings(MuingII_Settings.MuingDebug) = 1 Then
        If Alpha = 0 Then Alpha = 128
        Call DrawDebug
    End If

    If Alpha <> 0 Then Draw.Present GameDraw.CompatibleDC, 0, 0, Alpha
End Sub
Sub DrawDebug()
    Dim W As Long
    W = Int(GWW / UBound(HistoryDraw))
    For i = 0 To UBound(HistoryFPS)
        BackImg.ImageByIndex(7).PresentWithClip Draw.CompatibleDC, i * W, 0, 0, 0, W, HistoryFPS(i) * 2, HistoryFPS(i) / 100 * 255
        GameFont.DrawText Draw.Graphics, i * W, HistoryFPS(i) * 2, W, 30, HistoryFPS(i), argb(255, 255, 255, 255), center, 18, Bold
        GameFont.DrawText Draw.Graphics, i * W, HistoryFPS(i) * 2 + 20, W, 30, HistoryDraw(i), argb(255, 255, 255, 255), center, 12, Bold
    Next
    GameFont.DrawText Draw.Graphics, 20, 20, GWW, 30, HistoryFPS(UBound(HistoryFPS)) & " fps", argb(255, 255, 255, 255), near, 18, Bold
    GameFont.DrawText Draw.Graphics, 20, 50, GWW, 30, HistoryDraw(UBound(HistoryDraw)) & "x", argb(255, 255, 255, 255), near, 18, Bold
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    Back.Create TargetDC, GWW, GWH
    PageList(0) = "SoloPage": PageList(1) = "MultiPage": PageList(2) = "EditPage": PageList(3) = "DownPage": PageList(4) = "SetPage"
    AlphaId = Ani.NewTask(200, 100, 0, 1)
End Sub
