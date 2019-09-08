VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlayPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Public OpenTime As Long
Dim SLock As Boolean, DLock As Boolean, JLock As Boolean, KLock As Boolean, RePlay As Boolean
Dim GamePaused As Boolean, MenuBack As New Images
Dim FinishMark As Boolean, FinishIndex As Integer, HighScore As Boolean
Dim WinTime As Long
Dim MapC As Boolean
Dim RankStr As String
Dim MapOld As Boolean, MapChecked As Boolean
Dim ExpPlus As Long
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "PlayPage")
End Property

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "PlayPage"
    
    If MapC = False And Connected And UserID <> 0 And Newest Then
        If GetTickCount - OpenTime > 6000 Then OpenTime = GetTickCount - 5000
    End If
    
    If GetTickCount - OpenTime <= 6000 Then
        Call DrawPage0
        If GetTickCount - OpenTime <= 1500 Then MapChecked = False: MapC = False: MapCheckOK = False
        With MainImg.Image("backbutton.png")
            .SetClickArea GWW - .Width - 30, 30
            .Present Draw.CompatibleDC, GWW - .Width - 30, 30, IIf(IsMouseIn, 255, 200)
        End With
        If IsClick = True Then
            GameCore.ChangePage "SoloPage"
        End If
        GamePaused = False
        NowCombo = 0
        Score = 0: Combo = 0: SLock = False: DLock = False: JLock = False: KLock = False: RePlay = False
        ExcellentCount = 0: GoodCount = 0: OKCount = 0: MissCount = 0
        LastHit = 0: ActiveHit = 0: HitLate = 0: HitDownLate = 0
        PlayTime = 0: HP = 0
        HighScore = False
        FinishIndex = 0: FinishMark = False
        If GetTickCount - OpenTime > 2000 Then
            If Connected And UserID <> 0 And Newest And MapChecked = False Then
                MapC = False: MapCheckOK = False
                MapCheck = ""
                Send "checkmap*" & SongList(NowSong).Info.Artist & "*" & SongList(NowSong).Info.Maker & "*" & SongList(NowSong).Info.Title & "*" & SongList(NowSong).Info.MapID
                MapChecked = True
            End If
        End If
    ElseIf GamePaused = False Then
        Call DrawPage1
    Else
        Call DrawPage2
    End If
    
    Draw.Present GameDraw.CompatibleDC, 0, 0
End Sub
Public Sub DrawPage0()
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, 100
    
    SongList(NowSong).SongPic.Present Draw.CompatibleDC, 155, 200
    
    If MapC = False And Connected And UserID <> 0 And Newest Then
        BackImg.ImageByIndex(3).PresentWithClip Draw.CompatibleDC, 155, 200, 0, 0, 575, 104, 150
        GameFont.DrawText Draw.Graphics, 155, 240, 575, 30, "正在验证谱子", argb(255, 255, 255, 255), center, 18, Bold
    End If
    
    If MapCheck <> "" And MapC = False Then
        Dim temp() As String
        temp = Split(MapCheck, "*")
        MapOld = False
        If temp(1) <> "yes" Then
            If Val(temp(1)) > SongList(NowSong).Info.MapVersion + 1 Then
                MapOld = True
            Else
                MapCheckOK = True
                MapCheckID = Val(temp(3))
            End If
        End If
        MapC = True
    End If
    
    Dim Alpha As Long, Text(2) As String, Time As Long
    Time = GetTickCount
    Text(0) = "Easy": Text(1) = "Normal": Text(2) = "Hard"
    
    Alpha = 0
    If GetTickCount - OpenTime > 500 Then Alpha = IIf((Time - OpenTime) - 500 <= 1000, 255 * ((Time - OpenTime) - 500) / 1000, 255)
    GameFont.DrawText Draw.Graphics, 0, 320, GWW, 35, SongList(NowSong).Info.Title, argb(Alpha, 255, 255, 255), center, 24, Regular
    
    Alpha = 0
    If GetTickCount - OpenTime > 1500 Then Alpha = IIf((Time - OpenTime) - 1500 <= 1000, 255 * ((Time - OpenTime) - 1500) / 1000, 255)
    GameFont.DrawText Draw.Graphics, 0, 350, GWW, 35, Text(NowSongD), argb(Alpha, 255, 255, 255), center, 24, Bold
    
    Alpha = 0
    If GetTickCount - OpenTime > 2000 Then Alpha = IIf((Time - OpenTime) - 2000 <= 500, 150 * ((Time - OpenTime) - 2000) / 500, 150)
    GameFont.DrawText Draw.Graphics, 0, 400, GWW, 25, "Maker    " & SongList(NowSong).Info.Maker, argb(Alpha, 255, 255, 255), center, 16, Regular
    
    Alpha = 0
    If GetTickCount - OpenTime > 2000 Then Alpha = IIf((Time - OpenTime) - 2000 <= 500, 150 * ((Time - OpenTime) - 2000) / 500, 150)
    GameFont.DrawText Draw.Graphics, 0, 425, GWW, 25, "Artist    " & SongList(NowSong).Info.Artist, argb(Alpha, 255, 255, 255), center, 16, Regular
    
    If Connected = False Then
        GameNotify.TipText = "没有连接到服务器，你的成绩不会被上传。"
    Else
        If Newest Then
            If UserID = 0 Then GameNotify.TipText = "没有登陆账号，你的成绩不会被上传。"
        Else
            GameNotify.TipText = "你的游戏不是最新版本，你的成绩不会被上传。"
        End If
    End If
End Sub
Public Sub DrawFinish()
    If FinishIndex = 0 Then
        
        BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, 120
        
        GameFont.DrawText Draw.Graphics, 0, 150, GWW, 60, "Failed", argb(255, 255, 0, 0), center, 48
        
        PlayImg.Image("restart.png").PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 262, 300, "", 0, 1, Regular, Button2
        If IsClick = True Then
            HP = 0
            MenuSnd2.Play
            Call ReStart
        End If
        
        If IsMouseIn Then GameFont.DrawText Draw.Graphics, 262, 300 + 146, 136, 30, "重新开始", argb(255, 255, 255, 255), center, 18
        
        PlayImg.Image("stop.png").PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 485, 300, "", 0, 1, Regular, Button2
        If IsClick = True Then
            MenuSnd.Play
            GameCore.ChangePage "SoloPage"
        End If
        If IsMouseIn Then GameFont.DrawText Draw.Graphics, 485, 300 + 146, 136, 30, "退出", argb(255, 255, 255, 255), center, 18
        
        Exit Sub
    End If
    
    Dim Width2 As Long, Height2 As Long, Deep As Long, Path As Long, Brush As Long, Ox As Long
    Dim y As Long, x As Long, TargetI As Integer
    
    BackImg.ImageByIndex(1).Present Draw.CompatibleDC, 0, 0
    PlayImg.Image("songback2.png").Present Draw.CompatibleDC, -14, -9
    MainBack.PresentWithClip Draw.CompatibleDC, 0, 0, 0, 0, GWW, 225
    BackImg.ImageByIndex(3).PresentWithClip Draw.CompatibleDC, 0, 0, 0, 0, GWW, 225, 127
    
    'PlayImg.Image("songback.png").Present Draw.CompatibleDC, 20, 20
    SongList(NowSong).SongPicCircle.Present Draw.CompatibleDC, 70, 225 / 2 - 128 / 2
    
    Dim LevelText(2) As String
     
    LevelText(0) = "Easy": LevelText(1) = "Normal": LevelText(2) = "Hard"
    
    GameFont.DrawText Draw.Graphics, 240, 49, GWW, 35, SongList(NowSong).Info.Title & " - " & LevelText(NowSongD), argb(255, 255, 255, 255), near, 24, Bold
    
    GameFont.DrawText Draw.Graphics, 240, 84, GWW, 30, "Maker: " & SongList(NowSong).Info.Maker & " / Artist: " & SongList(NowSong).Info.Artist, argb(180, 255, 255, 255), near, 20, Regular
    DrawMod3 Draw.CompatibleDC, Draw.Graphics, 240, 129
    
    If GetTickCount - WinTime >= 0 Then GameFont.DrawText Draw.Graphics, 70, 260, GWW, 40, IIf(GetTickCount - WinTime <= 500, Int((GetTickCount - WinTime) / 500 * Combo), Combo) & " Max Combo", argb(255, 44, 41, 30), near, 18, Regular

    GameFont.DrawText Draw.Graphics, 70, 310, 136, 30, "EXCELLENT", argb(255, 107, 10, 230), near, 18, Regular
    If GetTickCount - WinTime >= 0 Then GameFont.DrawText Draw.Graphics, 210, 310, GWW, 30, IIf(GetTickCount - WinTime <= 500, Int((GetTickCount - WinTime) / 500 * ExcellentCount), ExcellentCount) & "x", argb(255, 44, 41, 30), near, 18, Regular
    
    GameFont.DrawText Draw.Graphics, 70, 340, 136, 30, "GREAT", argb(255, 255, 0, 144), near, 18, Regular
    If GetTickCount - WinTime >= 0 Then GameFont.DrawText Draw.Graphics, 210, 340, GWW, 30, IIf(GetTickCount - WinTime <= 500, Int((GetTickCount - WinTime) / 500 * GoodCount), GoodCount) & "x", argb(255, 44, 41, 30), near, 18, Regular
    
    GameFont.DrawText Draw.Graphics, 70, 370, 136, 30, "OKAY", argb(255, 52, 195, 225), near, 18, Regular
    If GetTickCount - WinTime >= 0 Then GameFont.DrawText Draw.Graphics, 210, 370, GWW, 30, IIf(GetTickCount - WinTime <= 500, Int((GetTickCount - WinTime) / 500 * OKCount), OKCount) & "x", argb(255, 44, 41, 30), near, 18, Regular
    
    GameFont.DrawText Draw.Graphics, 70, 400, 136, 30, "MISS", argb(255, 44, 41, 30), near, 18, Regular
    If GetTickCount - WinTime >= 0 Then GameFont.DrawText Draw.Graphics, 210, 400, GWW, 30, IIf(GetTickCount - WinTime <= 500, Int((GetTickCount - WinTime) / 500 * MissCount), MissCount) & "x", argb(255, 44, 41, 30), near, 18, Regular
    
    If ExpPlus <> 0 Then
        GameFont.DrawText Draw.Graphics, 70, 430, 136, 30, "Exp", argb(255, 107, 10, 230), near, 18, Regular
        If GetTickCount - WinTime >= 0 Then GameFont.DrawText Draw.Graphics, 210, 430, GWW, 30, IIf(GetTickCount - WinTime <= 500, Int((GetTickCount - WinTime) / 500 * ExpPlus), ExpPlus), argb(255, 44, 41, 30), near, 18, Regular
    End If
    
    If GetTickCount - WinTime >= 2000 Then GameFont.DrawText Draw.Graphics, 0, 317, GWW - 20, 40, IIf(GetTickCount - WinTime - 1000 <= 500, Int((GetTickCount - WinTime - 2000) / 500 * Score), Int(Score)), argb(255, 44, 41, 30), far, 24, Bold
    If GetTickCount - WinTime >= 1000 Then GameFont.DrawText Draw.Graphics, 787, 257, GWW, 40, IIf(GetTickCount - WinTime - 2000 <= 500, Int((GetTickCount - WinTime - 1000) / 500 * Accuracy * 1000) / 10, Int(Accuracy * 1000) / 10) & "%", argb(255, 44, 41, 30), near, 24, Bold
    
    If GetTickCount - WinTime >= 3000 Then
        If HighScore Then
            PlayImg.Image("newrecord.png").Present Draw.CompatibleDC, GWW - 250, 45
        End If
        
        If Grade <> "-" Then PlayImg.Image("gradesmall" & Grade & ".png").Present Draw.CompatibleDC, 707, 240
    
        If Connected Then
            If MapCheckOK Then
                GameFont.DrawText Draw.Graphics, 0, 367, GWW - 20, 40, UserName & " (" & Int(Score) & ")   #" & SongList(NowSong).Rank(NowSongD), argb(255, 107, 10, 230), far, 18, Regular
                GameFont.DrawText Draw.Graphics, 0, 407, GWW - 20, 200, RankStr, argb(255, 44, 41, 30), far, 18, Regular
            Else
                GameFont.DrawText Draw.Graphics, 0, 367, GWW - 20, 40, "该谱子不是最新的", argb(255, 44, 41, 30), far, 18, Regular
            End If
        Else
            GameFont.DrawText Draw.Graphics, 0, 367, GWW - 20, 40, "没有连接到服务器", argb(255, 44, 41, 30), far, 18, Regular
        End If
        
        With PlayImg.Image("okbutton.png")
            .SetClickArea 70, GWH - 122
            .Present Draw.CompatibleDC, 70, GWH - 122, IIf(IsMouseIn, 255, 200)
            If IsClick Then MenuSnd2.Play: GameCore.ChangePage "SoloPage"
        End With
        With PlayImg.Image("restartbutton.png")
            .SetClickArea 170, GWH - 122
            .Present Draw.CompatibleDC, 170, GWH - 122, IIf(IsMouseIn, 255, 200)
            If IsClick Then
                HP = 0
                MenuSnd2.Play
                Call ReStart
            End If
        End With
        
    End If
    
    If GetTickCount - WinTime < 3000 Then
        SetClickArea2 0, 0, GWW, GWH
        If IsClick Then MenuSnd.Play: WinTime = GetTickCount - 3500
    End If
End Sub
Sub FinishGame()
    Dim MODPlus As Single
    MODPlus = 1
    If ModPower(1) Then MODPlus = MODPlus * 1.2
    If ModPower(2) Then MODPlus = MODPlus * 1.2
    If ModPower(3) Then MODPlus = MODPlus * 1.1

    Score = Score * MODPlus
    
    HighScore = (Score > Val(GameSave.RSave(SongList(NowSong).Info.Title & "." & SongList(NowSong).Info.Maker & ".Score" & NowSongD)))
    If ModPower(0) Then HighScore = False
    
    If HighScore Then
        Dim ModText As String
        If ModPower(1) Then ModText = ModText & "Qu "
        If ModPower(2) Then ModText = ModText & "Fa "
        If ModPower(3) Then ModText = ModText & "Dr "
        GameSave.WSave SongList(NowSong).Info.Title & "." & SongList(NowSong).Info.Maker & "." & SongList(NowSong).Info.MapID & ".Score" & NowSongD, Score
        GameSave.WSave SongList(NowSong).Info.Title & "." & SongList(NowSong).Info.Maker & "." & SongList(NowSong).Info.MapID & ".Grade" & NowSongD, Grade
        GameSave.WSave SongList(NowSong).Info.Title & "." & SongList(NowSong).Info.Maker & "." & SongList(NowSong).Info.MapID & ".Combo" & NowSongD, Combo
        GameSave.WSave SongList(NowSong).Info.Title & "." & SongList(NowSong).Info.Maker & "." & SongList(NowSong).Info.MapID & ".Accuracy" & NowSongD, Accuracy
        GameSave.WSave SongList(NowSong).Info.Title & "." & SongList(NowSong).Info.Maker & "." & SongList(NowSong).Info.MapID & ".MOD" & NowSongD, ModText
        SongList(NowSong).Score(NowSongD) = Score
        SongList(NowSong).Grade(NowSongD) = Grade
        SongList(NowSong).MaxCombo(NowSongD) = Combo
        SongList(NowSong).Accuracy(NowSongD) = Accuracy
        SongList(NowSong).MODs(NowSongD) = ModText
        
        ExpPlus = 0
        
        If MapCheckOK = True Then
            ExpPlus = Round((Accuracy * SongList(NowSong).Difficulty(NowSongD) * MODPlus * 10))
            If ExpPlus < 0 Then ExpPlus = 1
            MsgCheck = ""
            Send "score*" & MapCheckID & "*" & NowSongD & "*" & Combo & "*" & Grade & "*" & ModText & "*" & Score & "*" & Accuracy
            Send "addexp*" & ExpPlus
            Exp = Exp + ExpPlus
            Do While MsgCheck = ""
                DoEvents
            Loop
            Dim Ranks() As String, tR() As String, Rank As Long
            Dim Po As Integer, Ranks2(4) As Integer, Ranks2_S(4) As Single, BigRank As Single
            Rank = 1
            Ranks = Split(MsgCheck, "*")
            For i = 1 To UBound(Ranks) - 1
                tR = Split(Ranks(i), ";")
                If Val(tR(4)) > Score Then Rank = Rank + 1
            Next
            RankStr = "": BigRank = 999999999999#
            For Po = 0 To 4
                For i = 1 To UBound(Ranks) - 1
                    tR = Split(Ranks(i), ";")
                    If Val(tR(4)) > Ranks2_S(Po) And Val(tR(4)) < BigRank Then
                        Ranks2_S(Po) = Val(tR(4))
                        Ranks2(Po) = i
                    End If
                Next
                If Ranks2(Po) = 0 Then
                    RankStr = RankStr & "空位 (0) #" & Po + 1 & vbCrLf
                Else
                    tR = Split(Ranks(Ranks2(Po)), ";")
                    BigRank = Val(tR(4))
                    RankStr = RankStr & tR(0) & " (" & Int(Val(tR(4))) & ")   #" & Po + 1 & vbCrLf
                End If
            Next
            
            SongList(NowSong).Rank(NowSongD) = Rank
            GameSave.WSave SongList(NowSong).Info.Title & "." & SongList(NowSong).Info.Maker & "." & SongList(NowSong).Info.MapID & ".Rank" & NowSongD, Rank
        End If
    End If

    GameCore.FadePage "PlayPage"
    FinishMark = True: FinishIndex = 1
    GameBGM.Pos = 0
    GameBGM.Play
    WinSnd.Play
    
    WinTime = GetTickCount
End Sub
Sub BreakGame()
    GameCore.FadePage "PlayPage"
    FinishMark = True: FinishIndex = 0
    GameBGM.StopMusic
End Sub
Public Sub DrawPage1()
    If PlayTime = 0 Then
        If RePlay = False Then
            GameCore.FadePage "PlayPage"
            PlayTime = GetTickCount: GameBGM.StopMusic: RePlay = True
            GameBGM.Pos = 0
        Else
            HP = 100: GameBGM.Play: PlayTime = 1
        End If
    Else
        If PlayTime <> 1 Then
            If GetTickCount - PlayTime >= 3000 Then
                PlayTime = 0
            Else
                HP = (GetTickCount - PlayTime) / 3000 * 100
                If MapOld Then GameNotify.TipText = "你的谱子不是最新的，成绩不会被上传。"
            End If
        End If
    End If
    
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, 50
    
    If FinishMark Then Call DrawFinish: Exit Sub
    
    LevelDrawing Draw.CompatibleDC, Draw.Graphics, 0, 150, False
    
    With PlayImg.Image("hpbar.png")
        .PresentWithClip Draw.CompatibleDC, 0, 0, 0, 0, HP / 100 * GWW, .Height
    End With
    
    GameFont.DrawText Draw.Graphics, 20, 30, 110, 50, Int(GameBGM.Pos / GameBGM.max * 100) & "%", argb(255, 255, 255, 255), center, 36, Bold
    GameFont.DrawText Draw.Graphics, GWW - 530, 30, 440, 50, Int(Score), argb(255, 255, 255, 255), far, 36, Bold
    If Grade <> "-" Then
        PlayImg.Image("gradesmall" & Grade & ".png").Present Draw.CompatibleDC, GWW - 80, 20
    End If
    
    GameFont.DrawText Draw.Graphics, GWW - 530, 90, 510, 50, NowCombo & " Combo / " & Int(Accuracy * 1000) / 10 & "%", argb(130, 255, 255, 255), far, 28, Regular
    
    With PlayImg.Image("pausebutton.png")
        .SetClickArea 22, 80
        .Present Draw.CompatibleDC, 22, 80, IIf(IsMouseIn, 255, 220)
        If IsClick And GamePaused = False Then
            GameBGM.Pause
            BlurTo MenuBack.CompatibleDC, 40
            GamePaused = True
            GameCore.FadePage "PlayPage"
        End If
    End With
    
    DrawMod2 Draw.CompatibleDC, Draw.Graphics, GWW - 20, 130
    
    If DebugMode Then
        If GetAsyncKeyState(VK_F6) < 0 Then
            ExcellentCount = UBound(PlayingLevel.MapData) + 1: GoodCount = 0: OKCount = 0: MissCount = 0
            Accuracy = 1: Combo = ExcellentCount: LastHit = UBound(PlayingLevel.MapData) + 1: ActiveHit = LastHit - 1
            Score = 0
            For NowCombo = 1 To Combo - 1
                Score = Score + 300 * (NowCombo / 8)
            Next
            Grade = GetGrade
            GameBGM.Pos = GameBGM.max - 1
        End If
    End If
    
    If HP < 0 Then Call BreakGame
    If GameBGM.Pos / GameBGM.max = 1 Or UBound(PlayingLevel.MapData) = 0 Then Call FinishGame
End Sub
Public Sub DrawPage2()
    MenuBack.Present Draw.CompatibleDC, 0, 0
    BackImg.ImageByIndex(3).PresentWithClip Draw.CompatibleDC, 0, 165, 0, 0, GWW, 270, 40
    
    PlayImg.Image("continue.png").PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 180, 230, "", 0, 1, Regular, Button2
    If IsClick = True Then MenuSnd.Play: GameBGM.Continue: GamePaused = False
    If IsMouseIn Then GameFont.DrawText Draw.Graphics, 180, 230 + 146, 136, 30, "继续", argb(255, 255, 255, 255), center, 18
    
    PlayImg.Image("restart.png").PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 367, 230, "", 0, 1, Regular, Button2
    If IsClick = True Then
        MenuSnd2.Play
        Call ReStart
    End If
    
    If IsMouseIn Then GameFont.DrawText Draw.Graphics, 367, 230 + 146, 136, 30, "重新开始", argb(255, 255, 255, 255), center, 18
    
    PlayImg.Image("stop.png").PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 560, 230, "", 0, 1, Regular, Button2
    If IsClick = True Then
        MenuSnd.Play
        GameCore.ChangePage "SoloPage"
    End If
    If IsMouseIn Then GameFont.DrawText Draw.Graphics, 560, 230 + 146, 136, 30, "退出", argb(255, 255, 255, 255), center, 18

End Sub
Sub ReStart()
    PlayingLevel = SongList(NowSong).Info.Levels(NowSongD)
    Score = 0: Combo = 0: NowCombo = 0: SLock = False: DLock = False: JLock = False: KLock = False: RePlay = False
    ExcellentCount = 0: GoodCount = 0: OKCount = 0: MissCount = 0
    LastHit = 0: ActiveHit = 0: HitLate = 0: HitDownLate = 0
    PlayTime = 0
    GameBGM.Pos = 0: GameBGM.Continue: GamePaused = False
    FinishIndex = 0: FinishMark = False
End Sub
Sub KeyDown(Key As Integer)
    If GetTickCount - OpenTime <= 6000 Then Exit Sub
    If ModPower(0) Then Exit Sub

    HitDown = GetTickCount
    HitDownLate = Abs((PlayingLevel.MapData(ActiveHit).StartTime / 1000) - GameBGM.Pos)
    If Key = vbKeyS Then
        PressLock = 0
        Call HitObj2(0)
    End If
    If Key = vbKeyD Then
        PressLock = 1
        Call HitObj2(1)
    End If
    If Key = vbKeyJ Then
        PressLock = 2
        Call HitObj2(2)
    End If
    If Key = vbKeyK Then
        PressLock = 3
        Call HitObj2(3)
    End If
End Sub
Sub KeyUp(Key As Integer)
    If GetTickCount - OpenTime <= 6000 Then Exit Sub
    
    PressLock = -1
    
    If Key = vbKeySpace Then SpacePress = True
    If Key = vbKeyEscape And GamePaused = False Then
        GameBGM.Pause
        BlurTo MenuBack.CompatibleDC, 40
        GamePaused = True
        GameCore.FadePage "PlayPage"
    End If
    If Key = vbKeyR Then
        Call ReStart
    End If
    
    If ModPower(0) Then Exit Sub
    If Key = vbKeyS Then Call HitObj(0)
    If Key = vbKeyD Then Call HitObj(1)
    If Key = vbKeyJ Then Call HitObj(2)
    If Key = vbKeyK Then Call HitObj(3)
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    MenuBack.Create TargetDC, GWW, GWH
End Sub



