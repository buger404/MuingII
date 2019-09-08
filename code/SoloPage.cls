VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoloPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images, Draw2 As New Images, Draw3 As New Images, Back As New Images
Dim ScrollY As Long, TargetY As Long, ScrollTime As Currency
Public ChangeTime As Currency, ChangeIndex As Integer
Dim RankMode As Boolean
Dim Ranks(4) As String
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "SoloPage")
End Property
Sub MouseWheel(Direction As Integer, Strong As Single)
    TargetY = IIf(GetTickCount - ScrollTime <= 100, TargetY, ScrollY) + 20 / UBound(SongList) * Strong * IIf(Direction = 1, 1, -1)
    ControlLong TargetY, 0, 100
    ScrollTime = GetTick
End Sub
Sub ScrollTo(ByVal SongIndex As Integer)

    '设ScrollY = x
    '40 - x / 100 * (UBound(SongList) * 134 + 170) + 110 + 134 * SongIndex = GWH / 2 - Height / 2
    '- x / 100 * (UBound(SongList) * 134 + 170)  = GWH / 2 - Height / 2 - 40 - 110 - 134 * SongIndex
    '- x / 100 = (GWH / 2 - Height / 2 - 40 - 110 - 134 * SongIndex) / (UBound(SongList) * 134 + 170)
    '- x = (GWH / 2 - Height / 2 - 40 - 110 - 134 * SongIndex) / (UBound(SongList) * 134 + 170) * 100
    'x = -((GWH / 2 - Height / 2 - 40 - 110 - 134 * SongIndex) / (UBound(SongList) * 134 + 170) * 100)
    
    Dim Height As Long
    Height = MainImg.Image("normalsong.png").Height
    
    TargetY = -((GWH / 2 - Height / 2 - 40 - 140 - 134 * (SongIndex - 1)) / (UBound(SongList) * 134 + 100) * 100): ScrollTime = GetTick
    ControlLong TargetY, 0, 100
End Sub
Sub DrawRank()
    Back.Present Draw.CompatibleDC, 0, 0
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, HotAlpha
    
    Dim DrawX As Long, LevelText(2) As String
    LevelText(0) = "Easy": LevelText(1) = "Normal": LevelText(2) = "Hard"
     
    With MainImg.Image("rankframe.png")
        DrawX = GWW / 2 - .Width / 2
        .Present Draw.CompatibleDC, DrawX, 70
    End With
    
    SongList(NowSong).SongPic.Present Draw.CompatibleDC, DrawX + 13, 81, 40
    SongList(NowSong).SongPicCircle.Present Draw.CompatibleDC, DrawX + 46, 110
    
    GameFont.DrawText Draw.Graphics, DrawX, 105, 560, 30, SongList(NowSong).Info.Title, argb(255, 255, 255, 255), far, 20, Bold
    GameFont.DrawText Draw.Graphics, DrawX, 135, 560, 30, LevelText(NowSongD), argb(255, 255, 255, 255), far, 20
    
    Dim tR() As String, DrawX2 As Long
    
    If Ranks(0) = "" Then
        GameFont.DrawText Draw.Graphics, DrawX, 325, 602, 30, "加载中...", argb(255, 27, 27, 27), center, 16
    Else
        If Ranks(0) = "空" Then
            GameFont.DrawText Draw.Graphics, DrawX, 325, 602, 30, "还没有人玩这张谱子哦", argb(255, 27, 27, 27), center, 16
        End If
        For i = 0 To 3
            If Ranks(i) <> "空" Then
                tR = Split(Ranks(i), ";")
                PlayImg.Image("gradesmall" & tR(2) & ".png").Present Draw.CompatibleDC, DrawX + 50, 240 + i * 60
                GameFont.DrawText Draw.Graphics, DrawX + 120, 240 + i * 60 + 12, 300, 30, tR(0), argb(255, 27, 27, 27), near, 18, Regular
                GameFont.DrawText Draw.Graphics, DrawX + 120, 240 + i * 60 + 32, 300, 30, Int(tR(4)) & "  " & tR(1) & " max combo / " & Int(Val(tR(5)) * 1000) / 10 & "%", argb(200, 27, 27, 27), near, 16, Regular
                If tR(3) <> "" Then
                    Dim ModP(3) As Boolean
                    If InStr(tR(3), "Qu") Then ModP(1) = True
                    If InStr(tR(3), "Fa") Then ModP(2) = True
                    If InStr(tR(3), "Dr") Then ModP(3) = True
                    DrawX2 = DrawX + 510
                    For s = 3 To 0 Step -1
                        If ModP(s) Then
                            With ModImg.Image("mod" & s & "active.png")
                                .Present Draw.CompatibleDC, DrawX2, 240 + i * 60 + 7
                                DrawX2 = DrawX2 - .Width - 3
                            End With
                        End If
                    Next
                End If
            End If
        Next
    End If
    
    If MsgCheck <> "" Then
        Dim Rank() As String
        Dim Po As Integer, Rank2(4) As Integer, Rank2_S(4) As Single, BigRank As Single
        Rank = Split(MsgCheck, "*")
        BigRank = 999999999999#
        For Po = 0 To 4
            For i = 1 To UBound(Rank) - 1
                tR = Split(Rank(i), ";")
                If Val(tR(4)) > Rank2_S(Po) And Val(tR(4)) < BigRank Then
                    Rank2_S(Po) = Val(tR(4))
                    Rank2(Po) = i
                End If
            Next
            If Rank2(Po) = 0 Then
                Ranks(Po) = "空"
            Else
                Ranks(Po) = Rank(Rank2(Po))
                tR = Split(Ranks(Po), ";")
                BigRank = Val(tR(4))
            End If
        Next
        MsgCheck = ""
    End If
    
    If Ranks(0) <> "" Then
        CtrlImg.ImageByIndex(1).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, DrawX + 440, 475, "OK", argb(255, 255, 255, 255), 16, Regular, Button
        If IsClick Then
            MenuSnd2.Play: RankMode = False
            GameCore.FadePage "SoloPage"
            KeepSong = False
        End If
    End If
End Sub
Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "SoloPage"
    
    If RankMode Then
        NowPage = "RankMode"
        Call DrawRank
        Draw.Present GameDraw.CompatibleDC, 0, 0
        Exit Sub
    End If
    
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, HotAlpha
    
    If NowSong = 0 Then
        GameFont.DrawText Draw.Graphics, 0, GWH / 2 - 30 / 2, GWW, 30, "没有谱子，快去下载吧！", argb(255, 27, 27, 27), center, 16
    End If
    
    If GetPastTick(ScrollTime) <= 100 Then
        ScrollY = GetPastTick(ScrollTime) / 100 * (TargetY - ScrollY) + ScrollY
    ElseIf ScrollTime <> 0 Then
        ScrollY = TargetY
        ScrollTime = 0
    End If
    
    Dim DrawY As Long
    Dim Frame1 As Images, Frame2 As Images, Shadow1 As Images, Shadow2 As Images
    Dim Points(2) As Images
    Dim ChangePro As Single, EffectX As Long
    Set Frame1 = MainImg.Image("focusframe.png"): Set Frame2 = MainImg.Image("normalsong.png")
    Set Shadow1 = MainImg.Image("focusshadow.png"): Set Shadow2 = MainImg.Image("normalshadow.png")
    Set Points(0) = MainImg.Image("easypoint.png"): Set Points(1) = MainImg.Image("normalpoint.png"): Set Points(2) = MainImg.Image("hardpoint.png")
    DrawY = 40 - ScrollY / 100 * (UBound(SongList) * 134 + 50)
    'MainImg.Image("solo.png").Present Draw.CompatibleDC, 40, DrawY
    DrawY = DrawY + 40

    '40 - ScrollY / 100 * (UBound(SongList) * 134 + 170) = 200
    'ScrollY  = -(200 / (UBound(SongList) * 134 + 170) * 100 -40)
    
    If GetPastTick(ChangeTime) <= 300 Then ChangePro = cubicCurves(GetPastTick(ChangeTime) / 500, 0, 1, 1, 1)

    Dim DrawDone() As Boolean

    For i = 1 To UBound(SongList)
        If DrawY >= IIf(NowSong = i, -134 - 160, -134) And DrawY <= GWH Then
            EffectX = Sin(DrawY / 90) * 10
            
            If NowSong = i Then
                Frame1.Present Draw.CompatibleDC, 40 + EffectX, DrawY
            Else
                Frame2.Present Draw.CompatibleDC, 40 + EffectX, DrawY
            End If
            SongList(i).SongPic.Present Draw.CompatibleDC, 55 + EffectX, DrawY + 15
            
            If NowSong = i Then Shadow1.Present Draw.CompatibleDC, 40 + EffectX, DrawY

            SetClickArea2 55 + EffectX, DrawY + 15, 575, 102
            If IsClick = True And NowSong <> i Then
                MenuSnd.Play
                SetMainBack SongPath & "\" & SongList(i).Path & "\background.png"
                ChangeIndex = NowSong: ChangeTime = GetTick
                GdipGraphicsClear Draw3.Graphics, 0
                Draw2.Present Draw3.CompatibleDC, 0, 0
                NowSong = i
                GameBGM.LoadMusic SongPath & "\" & SongList(i).Path & "\music.mp3"
                GameBGM.Play
                ScrollTo i
            End If
            
            GameFont.DrawText Draw.Graphics, 75 + EffectX, DrawY + 32, 534, 30, SongList(i).Info.Title, argb(255, 255, 255, 255), near, 16, Bold
            GameFont.DrawText Draw.Graphics, 75 + EffectX, DrawY + 52, 534, 30, SongList(i).Info.Maker & " / " & SongList(i).Info.Artist, argb(255, 255, 255, 255), near, 14
     
            If NowSong = i Then
                Dim SongLen As Long, SongLenT As String, CAble As Boolean
                
                SongLen = GameBGM.max * IIf(ModPower(1), 0.75, 1)
                SongLenT = Int(SongLen / 60) & ":" & format(SongLen Mod 60, "00")
                
                With MainImg.Image("edit.png")
                    .SetClickArea 660 + EffectX, DrawY + 27
                    .Present Draw.CompatibleDC, 660 + EffectX, DrawY + 27, IIf(IsMouseIn, 255, 150)
                    If IsClick = True Then
                        Open SongPath & "\" & SongList(i).Path & "\map.mu" For Binary As #1
                        Get #1, , EditMap
                        Close #1
                        PlayingLevel = EditMap.Levels(Mode)
                        SetMainBack SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\Background.png"
                        SLock = False: DLock = False: JLock = False: KLock = False: Saved = False: Mode = 0
                        EditPage.ReDrawHitPic
                        EditPage.Page = 1
                        NowSong = i
                        EditPage.Saved = True
                        GameCore.FadePage "EditPage"
                        mNowShow = "EditPage"
                        GameBGM.LoadMusic SongPath & "\" & SongList(i).Path & "\music.mp3"
                        GameBGM.StopMusic
                    End If
                End With
                With MainImg.Image("delete.png")
                    .SetClickArea 760 + EffectX, DrawY + 27
                    .Present Draw.CompatibleDC, 760 + EffectX, DrawY + 27, IIf(IsMouseIn, 255, 150)
                    If IsClick Then
                        If GameNotify.Asks(MuingAsk, "即将删除谱子，此操作不可逆，继续吗？", "点错了", "删除") = 0 Then Exit Sub
                        GameBGM.StopMusic
                        GameBGM.Dispose
                        Dim FSO As Object
                        Set FSO = CreateObject("Scripting.FileSystemObject")
                        FSO.DeleteFolder SongPath & "\" & SongList(i).Path
                        GetSongList
                        Exit Sub
                    End If
                End With
                MainImg.Image("time.png").Present Draw.CompatibleDC, 430 + EffectX, DrawY + 33
                MainImg.Image("object0.png").Present Draw.CompatibleDC, 530 + EffectX, DrawY + 33
                
                GameFont.DrawText Draw.Graphics, 430 + EffectX + 25, DrawY + 35, 70, 30, SongLenT, argb(255, 255, 255, 255), center, 14, Bold
                GameFont.DrawText Draw.Graphics, 530 + EffectX + 25, DrawY + 35, 70, 30, UBound(SongList(NowSong).Info.Levels(NowSongD).MapData) & "x", argb(255, 255, 255, 255), center, 14, Bold
                
                If Connected And Newest Then
                    With MainImg.Image("rankbutton.png")
                        .SetClickArea 570 + EffectX, DrawY + 60
                        .Present Draw.CompatibleDC, 570 + EffectX, DrawY + 60, IIf(IsMouseIn, 255, 200)
                        If IsMouseUp Then
                            KeepSong = True
                            MsgCheck = "": Ranks(0) = ""
                            Send "getrank*" & SongList(NowSong).Info.Maker & "*" & SongList(NowSong).Info.Title & "*" & SongList(NowSong).Info.Artist & "*" & SongList(NowSong).Info.MapID & "*" & NowSongD
                            BlurTo Back.CompatibleDC, 60
                            MenuSnd2.Play
                            RankMode = True
                            GameCore.FadePage "SoloPage"
                        End If
                    End With
                End If
            End If
            
            Dim MaxCount As Long
            MaxCount = 0
            ReDim DrawDone(18)
            
            If NowSongD = 0 Or i <> NowSong Then
                For s = 1 To Int(SongList(i).Difficulty(0))
                    If s > 18 Then
                        If Int(SongList(i).Difficulty(0)) > MaxCount Then MaxCount = Int(SongList(i).Difficulty(0))
                        Exit For
                    End If
                    If DrawDone(s) = False Then Points(0).Present Draw.CompatibleDC, 78 + 28 * (s - 1) + EffectX, DrawY + 78: DrawDone(s) = True
                Next
            End If
            If NowSongD = 1 Or i <> NowSong Then
                For s = 1 To Int(SongList(i).Difficulty(1))
                    If s > 18 Then
                        If Int(SongList(i).Difficulty(1)) > MaxCount Then MaxCount = Int(SongList(i).Difficulty(1))
                        Exit For
                    End If
                    If DrawDone(s) = False Then Points(1).Present Draw.CompatibleDC, 78 + 28 * (s - 1) + EffectX, DrawY + 78: DrawDone(s) = True
                Next
            End If
            If NowSongD = 2 Or i <> NowSong Then
                For s = 1 To Int(SongList(i).Difficulty(2))
                    If s > 18 Then
                        If Int(SongList(i).Difficulty(2)) > MaxCount Then MaxCount = Int(SongList(i).Difficulty(2))
                        Exit For
                    End If
                    If DrawDone(s) = False Then Points(2).Present Draw.CompatibleDC, 78 + 28 * (s - 1) + EffectX, DrawY + 78: DrawDone(s) = True
                Next
            End If
            
            If MaxCount <> 0 Then GameFont.DrawText Draw.Graphics, 70 + 28 * 18 + EffectX, DrawY + 78, 53, 30, "+" & (MaxCount - 18), argb(255, 255, 255, 255), center, 18, Bold
            
            If NowSong = i Then
                GdipGraphicsClear Draw2.Graphics, 0
                BackImg.ImageByIndex(3).PresentWithClip Draw2.CompatibleDC, 0, DrawY + 134 - (DrawY + 134), 0, 0, GWW, 160, 51
                CAble = (UBound(SongList(i).Info.Levels(0).MapData) > 0)
                MainImg.Image("easyframe" & IIf((NowSongD = 0) And CAble, "2", "") & ".png").Present Draw2.CompatibleDC, 70, DrawY + 134 - (DrawY + 134), IIf(CAble, 255, 100)
                If CAble Then SetClickArea2 70, DrawY + 134, 581, 61
                If IsClick = True And CAble Then
                    If NowSongD = 0 Then
                        PlayingLevel = SongList(i).Info.Levels(NowSongD)
                        PlayPage.OpenTime = GetTickCount
                        MenuSnd2.Play
                        GameCore.ChangePage "PlayPage"
                    Else
                        MenuSnd.Play
                        NowSongD = 0
                    End If
                End If
                GameFont.DrawText Draw2.Graphics, 130, DrawY + 134 + 19 - (DrawY + 134), 300, 20, "Easy" & IIf(CAble, "", " (No Object)"), argb(IIf(CAble, 255, 100), 255, 255, 255), near, 16, Bold
                GameFont.DrawText Draw2.Graphics, 205, DrawY + 134 + 19 - (DrawY + 134), 415, 20, "#" & SongList(i).Rank(0) & "      " & SongList(i).Grade(0), argb(IIf(CAble, 255, 100), 255, 255, 255), far, 18, Bold
                
                CAble = (UBound(SongList(i).Info.Levels(1).MapData) > 0)
                MainImg.Image("normalframe" & IIf((NowSongD = 1) And CAble, "2", "") & ".png").Present Draw2.CompatibleDC, 70, DrawY + 134 + 48 - (DrawY + 134), IIf(CAble, 255, 100)
                If CAble Then SetClickArea2 70, DrawY + 134 + 48, 581, 61
                If IsClick = True And CAble Then
                    If NowSongD = 1 Then
                        PlayingLevel = SongList(i).Info.Levels(NowSongD)
                        PlayPage.OpenTime = GetTickCount
                        MenuSnd2.Play
                        GameCore.ChangePage "PlayPage"
                    Else
                        MenuSnd.Play
                        NowSongD = 1
                    End If
                End If
                GameFont.DrawText Draw2.Graphics, 130, DrawY + 134 + 48 + 19 - (DrawY + 134), 300, 20, "Normal" & IIf(CAble, "", " (No Object)"), argb(IIf(CAble, 255, 100), 255, 255, 255), near, 16, Bold
                GameFont.DrawText Draw2.Graphics, 205, DrawY + 134 + 48 + 19 - (DrawY + 134), 415, 20, "#" & SongList(i).Rank(1) & "      " & SongList(i).Grade(1), argb(IIf(CAble, 255, 100), 255, 255, 255), far, 18, Bold
                
                CAble = (UBound(SongList(i).Info.Levels(2).MapData) > 0)
                MainImg.Image("hardframe" & IIf((NowSongD = 2) And CAble, "2", "") & ".png").Present Draw2.CompatibleDC, 70, DrawY + 134 + 48 * 2 - (DrawY + 134), IIf(CAble, 255, 100)
                If CAble Then SetClickArea2 70, DrawY + 134 + 48 * 2, 581, 61
                If IsClick = True And CAble Then
                    If NowSongD = 2 Then
                        PlayingLevel = SongList(i).Info.Levels(NowSongD)
                        PlayPage.OpenTime = GetTickCount
                        MenuSnd2.Play
                        GameCore.ChangePage "PlayPage"
                    Else
                        MenuSnd.Play
                        NowSongD = 2
                    End If
                End If
                GameFont.DrawText Draw2.Graphics, 130, DrawY + 134 + 48 * 2 + 19 - (DrawY + 134), 300, 20, "Hard" & IIf(CAble, "", " (No Object)"), argb(IIf(CAble, 255, 100), 255, 255, 255), near, 16, Bold
                GameFont.DrawText Draw2.Graphics, 205, DrawY + 134 + 48 * 2 + 19 - (DrawY + 134), 415, 20, "#" & SongList(i).Rank(2) & "      " & SongList(i).Grade(2), argb(IIf(CAble, 255, 100), 255, 255, 255), far, 18, Bold
                GameFont.DrawText Draw2.Graphics, GWW - 200 - 25, DrawY + 134 + 15 - (DrawY + 134), 190, 50, SongList(i).Score(NowSongD) & " " & SongList(i).MaxCombo(NowSongD) & "x", argb(255, 255, 255, 255), center, 22, Bold
                GameFont.DrawText Draw2.Graphics, GWW - 200 - 25, DrawY + 134 + 45 - (DrawY + 134), 190, 50, SongList(i).MODs(NowSongD) & Int(SongList(i).Accuracy(NowSongD) * 1000) / 10 & "%", argb(200, 255, 255, 255), far, 22, Regular
                
                SetClickArea2 GWW - 200 - 35, DrawY + 134 + 85, 214, 58
                DrawMod Draw2.CompatibleDC, Draw2.Graphics, GWW - 200 - 27, DrawY + 134 + 89
                ModImg.Image("modframe.png").Present Draw2.CompatibleDC, GWW - 200 - 35, DrawY + 134 + 85 - (DrawY + 134), 255
                'If IsClick = True Then GameNotify.Message "- ^ O", "This is a MOD box , what's up ?", "oh♂"
            Else
                Shadow2.Present Draw.CompatibleDC, 40 + EffectX, DrawY, Abs(DrawY - GWH / 2) / (GWH / 2) * 120 + 135
            End If
        End If
        If i = NowSong Then
            If ChangePro <> 0 Then
                Draw2.PresentWithClip Draw.CompatibleDC, 0, DrawY + 134, 0, 0, GWW, ChangePro * 160, ChangePro * 255
                DrawY = DrawY + ChangePro * 160
            Else
                Draw2.Present Draw.CompatibleDC, 0, DrawY + 134
                DrawY = DrawY + 160
            End If
        End If
        
        If i = ChangeIndex Then
            If ChangePro <> 0 Then
                Draw3.PresentWithClip Draw.CompatibleDC, 0, DrawY + 134, 0, 0, GWW, 160 - ChangePro * 160, 255 - ChangePro * 255
                DrawY = DrawY + (160 - ChangePro * 160)
            End If
        End If
        
        If DrawY > GWH Then Exit For
        DrawY = DrawY + 134
    Next
    
    CtrlImg.ImageByIndex(BMCtrl.BMVScroll).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 860, 100, ScrollY, 0, 16, Regular, VScrollBar2, , 100

    SetClickArea2 GWW - 210, -20, 210, 146
    'MainImg.Image("backbutton.png").Present Draw.CompatibleDC, GWW - 210, -20, IIf(IsMouseIn, 255, 170)
    If IsClick = True Then
        GameCore.ChangePage "MainPage"
    End If
    
    Draw.Present GameDraw.CompatibleDC, 0, 0
End Sub

Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    Back.Create TargetDC, GWW, GWH
    Draw2.Create TargetDC, GWW, 160: Draw3.Create TargetDC, GWW, 160
End Sub



