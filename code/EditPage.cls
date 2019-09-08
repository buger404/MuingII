VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images, HitPic As New Images
Public Page As Integer, Saved As Boolean, Mode As Integer
Dim SLock As Boolean, DLock As Boolean, JLock As Boolean, KLock As Boolean
Dim ScrollY As Long
Dim UpdateTime As Long
Dim PressDown As Long
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "EditPage")
End Property

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "EditPage"
    
    If Page = 0 Then
        Call DrawPage0
    Else
        Call DrawPage1
    End If
    
    Draw.Present GameDraw.CompatibleDC, 0, 0
End Sub
Sub FileDrop(Path As String)
    If Page = 0 Then
        If CheckFileHeader(Path, "ID3") = False Then
            GameNotify.Message MuingError, "你拖进来的不是.mp3文件哟~", "soga"
        Else
            If Connected = False Then GameNotify.Message MuingNoConnected, "很抱歉，我们需要连接服务器来检查你的谱子ID是否被占用。", "噢": Exit Sub
            If Newest = False Then GameNotify.Message MuingUpdate, "很抱歉，你的游戏版本不是最新的。", "噢": Exit Sub
            Page = 1
            GdipGraphicsClear MainBack.Graphics, argb(255, 249, 249, 249)
            Dim temp As MuingII_MapFile
            temp.MuingVersion = 1
            temp.Levels(0).ObjSpeed = 0.8: temp.Levels(1).ObjSpeed = 1: temp.Levels(2).ObjSpeed = 1.5
            ReDim temp.Levels(0).MapData(0)
            ReDim temp.Levels(1).MapData(0)
            ReDim temp.Levels(2).MapData(0)
            SLock = False: DLock = False: JLock = False: KLock = False
            EditMap = temp
            SongCreateWindow.Show
            SongCreateWindow.IDText.Enabled = True
            SongCreateWindow.ArtistText.Enabled = True: SongCreateWindow.MakerText.Enabled = True: SongCreateWindow.MusicText.Enabled = True
            Saved = False: Mode = 0
            Do While MainWindow.Enabled = False
                DoEvents
            Loop
            PlayingLevel = EditMap.Levels(Mode)
            FileCopy Path, SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\music.mp3"
            GameBGM.LoadMusic SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\music.mp3"
            GameBGM.StopMusic
        End If
    Else
        If CheckFileHeader(Path, Chr(137) & Chr(80) & "NG") = False Then
            GameNotify.Message MuingError, "我们只能支持.png作为背景图片哦~", "噢。"
        Else
            FileCopy Path, SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\Background.png"
            SetMainBack SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\Background.png"
            'BackImg.ImageByIndex(1).Present MainBack.CompatibleDC, 0, 0, 50
        End If
    End If
End Sub
Public Sub DrawPage0()
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    EditImg.Image("drag.png").Present Draw.CompatibleDC, 0, 0
End Sub
Public Sub DrawPage1()
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, 120
    
    HP = 100
    
    Dim Color As Long, ModeText As String
    If Mode = 0 Then Color = argb(255, 255, 192, 0): ModeText = "Easy"
    If Mode = 1 Then Color = argb(255, 0, 174, 245): ModeText = "Normal"
    If Mode = 2 Then Color = argb(255, 255, 2, 141): ModeText = "Hard"

    GameFont.DrawText Draw.Graphics, 30, 83, GWW, 25, EditMap.Title & " - " & ModeText & "  " & GetDifficulty(EditMap.Levels(Mode)) & "*", Color, near, 16
    
    LevelDrawing Draw.CompatibleDC, Draw.Graphics, 0, 160, True
    
    GameFont.DrawText Draw.Graphics, 0, 83, 650, 25, Int(GameBGM.Pos / 60) & ":" & format((Int(GameBGM.Pos * 10) Mod 600) / 10, "00.00") & " / " & Int(GameBGM.max / 60) & ":" & format((Int(GameBGM.max * 10) Mod 600) / 10, "00.00"), argb(255, 255, 255, 255), far, 16
    'GameFont.DrawText Draw.Graphics, 10, GWH - 30, 600, 25, "Tips: S/D/J/K 击打 , S/D/J/K + L 长条 , 选中物件后按Del可以删除物件。", argb(255, VC1, VC2, VC3), Near, 16

    HitPic.Present Draw.CompatibleDC, 30, 115, 180
    SetClickArea2 30, 115, 620, 40
    If IsMouseDown Or IsMouseUp Then
        If MouseX < 30 Then MouseX = 30
        If MouseX > 620 + 30 Then MouseX = 620 + 30
        GameBGM.Pos = (MouseX - 30) / 620 * GameBGM.max
        LastHit = 0
    End If

    EditImg.Image("playline.png").Present Draw.CompatibleDC, 30 + GameBGM.Pos / GameBGM.max * 620, 106

    With EditImg.Image("play.png")
        .Present Draw.CompatibleDC, 680, 115, IIf(GameBGM.PlayState = Playing, 120, 255)
        .SetClickArea 680, 115
        If IsClick = True Then GameBGM.Continue: LastHit = 0: GameCore.FadePage mNowShow
    End With
    
    With EditImg.Image("stop.png")
        .Present Draw.CompatibleDC, 730, 115, IIf(GameBGM.PlayState = Playing, 255, 120)
        .SetClickArea 730, 115
        If IsClick = True Then GameBGM.Pause: ReDrawHitPic: GameCore.FadePage mNowShow
    End With

    With EditImg.Image("erase.png")
        .SetClickArea 780, 115
        .Present Draw.CompatibleDC, 780, 115, IIf(IsMouseIn, 255, 120)
        If IsClick = True Then
            If GameNotify.Asks(MuingAsk, "擦除当前难度？此操作不可逆！", "取消", "擦除") = 0 Then Exit Sub
            ReDim EditMap.Levels(Mode).MapData(0)
            PlayingLevel = EditMap.Levels(Mode)
            GameBGM.Pause: ReDrawHitPic: GameCore.FadePage mNowShow
        End If
    End With
    
    With EditImg.Image("easy.png")
        .Present Draw.CompatibleDC, 820, 200
        .SetClickArea 820, 200
        If IsMouseUp = True Then
            Mode = 0: ReDrawHitPic: LastHit = 0: PlayingLevel = EditMap.Levels(Mode): GameCore.FadePage mNowShow
        End If
        If Mode = 0 Then EditImg.Image("choice.png").Present Draw.CompatibleDC, 820 - 2, 200 - 3
    End With
    With EditImg.Image("normal.png")
        .Present Draw.CompatibleDC, 820, 260
        .SetClickArea 820, 260
        If IsClick Then Mode = 1: ReDrawHitPic: LastHit = 0: PlayingLevel = EditMap.Levels(Mode): GameCore.FadePage mNowShow
        If Mode = 1 Then EditImg.Image("choice.png").Present Draw.CompatibleDC, 820 - 2, 260 - 3
    End With
    With EditImg.Image("hard.png")
        .Present Draw.CompatibleDC, 820, 320
        .SetClickArea 820, 320
        If IsClick Then Mode = 2: ReDrawHitPic: LastHit = 0: PlayingLevel = EditMap.Levels(Mode): GameCore.FadePage mNowShow
        If Mode = 2 Then EditImg.Image("choice.png").Present Draw.CompatibleDC, 820 - 2, 320 - 3
    End With
    
    With EditImg.Image("setting.png")
        .Present Draw.CompatibleDC, 820, 380
        .SetClickArea 820, 380
        If IsClick Then SongCreateWindow.Show: SongCreateWindow.IDText.Enabled = False: SongCreateWindow.ArtistText.Enabled = False: SongCreateWindow.MakerText.Enabled = False: SongCreateWindow.MusicText.Enabled = False: Saved = False
    End With
    
    
    If Saved = False Then
        'CtrlImg.ImageByIndex(BMCtrl.BMNormalButton).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 740, 530, "保存", argb(255, 255, 255, 255), 16, Regular, Button
        With EditImg.Image("save.png")
            .SetClickArea 75, GWH - 60
            .Present Draw.CompatibleDC, 75, GWH - 60, IIf(IsMouseIn, 255, 180)
        End With
        If IsClick = True Then
            EditMap.MapVersion = EditMap.MapVersion + 1
            Open SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\map.mu" For Binary As #1
            Put #1, , EditMap
            Close #1
            Saved = True
            GetSongList
        End If
    End If
    
    'CtrlImg.ImageByIndex(BMCtrl.BMNormalButton).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 740 - 150, 530, "导出", argb(255, 255, 255, 255), 16, Regular, Button
    With EditImg.Image("export.png")
        .SetClickArea 20, GWH - 60
        .Present Draw.CompatibleDC, 20, GWH - 60, IIf(IsMouseIn, 255, 180)
    End With
    If IsClick = True Then
        MakePackage App.Path & "\" & EditMap.Title & " - " & EditMap.Maker & ".mumap", SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\"
        GameNotify.Message MuingSuccess, "导出成功" & vbCrLf & App.Path & "\" & EditMap.Title & " - " & EditMap.Maker & ".mumap", "在哪啊"
    End If
    
    If Connected And (GetTickCount - UpdateTime) >= 5000 And Newest And UserID <> 0 Then
        'CtrlImg.ImageByIndex(BMCtrl.BMNormalButton).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 740, 470, "上传", argb(255, 255, 255, 255), 16, Regular, Button
        With EditImg.Image("upload.png")
            .SetClickArea 75 + 55, GWH - 60
            .Present Draw.CompatibleDC, 75 + 55, GWH - 60, IIf(IsMouseIn, 255, 180)
        End With
        If IsClick = True Then
            If Saved = False Then
                GameNotify.Asks MuingInformation, "请先保存您的谱子。", "哦"
                Exit Sub
            End If
            If EditMap.DownloadSource = "/" Then
                GameNotify.Asks MuingInformation, "请指定正确的下载地址！", "哦"
                SongCreateWindow.Show: Saved = False
                SongCreateWindow.IDText.Enabled = False
                SongCreateWindow.ArtistText.Enabled = False: SongCreateWindow.MakerText.Enabled = False: SongCreateWindow.MusicText.Enabled = False
                Exit Sub
            End If
            GameNotify.Asks MuingInformation, "提示：上传到服务器的只有下载地址。", "好的"
            UpdateTime = GetTickCount
            Send "updatemap*" & EditMap.Artist & "*" & EditMap.Maker & "*" & EditMap.Title & "*" & EditMap.MapID & "*" & EditMap.MapVersion & "*" & EditMap.DownloadSource
            GameNotify.Asks MuingSuccess, "上传成功！", "好"
        End If
    End If
End Sub
Sub KeyDown(Key As Integer)
    If Page = 0 Then Exit Sub

    If Key = vbKeyS Or Key = vbKeyD Or Key = vbKeyJ Or Key = vbKeyK Then
        HitSnd.Play
        PressDown = GameBGM.Pos * 1000
        'Debug.Print Now, "Down", PressDown
    End If
End Sub
Sub KeyUp(Key As Integer)
    If Page = 0 Then Exit Sub
    Saved = False
    Dim Pos As Long, Line As Long
    Pos = GameBGM.Pos * 1000
    
    If Key = vbKeySpace Then SpacePress = True
    If Key = vbKeyDelete Then
        If UBound(EditMap.Levels(Mode).MapData) = 0 Then Exit Sub
        EraseHitPic EditIndex
        For i = EditIndex To UBound(EditMap.Levels(Mode).MapData) - 1
            EditMap.Levels(Mode).MapData(i) = EditMap.Levels(Mode).MapData(i + 1)
        Next
        ReDim Preserve EditMap.Levels(Mode).MapData(UBound(EditMap.Levels(Mode).MapData) - 1)
        PlayingLevel = EditMap.Levels(Mode)
    End If
    
    If Key = vbKeyS Then Line = 0: GoTo SetObj0
    If Key = vbKeyD Then Line = 1: GoTo SetObj0
    If Key = vbKeyJ Then Line = 2: GoTo SetObj0
    If Key = vbKeyK Then Line = 3: GoTo SetObj0
        
    Exit Sub
    
SetObj0:
    Dim Index As Integer
    Dim Late As Long
    Late = (Pos - PressDown)
    'Debug.Print Now, "Up", Pos
    
    ReDim Preserve EditMap.Levels(Mode).MapData(UBound(EditMap.Levels(Mode).MapData) + 1)
    
    Index = UBound(EditMap.Levels(Mode).MapData)
    For i = 1 To UBound(EditMap.Levels(Mode).MapData) - 1
        If EditMap.Levels(Mode).MapData(i).StartTime > Pos Then Index = i: Exit For
    Next
    
    For i = UBound(EditMap.Levels(Mode).MapData) To Index + 1 Step -1
        EditMap.Levels(Mode).MapData(i) = EditMap.Levels(Mode).MapData(i - 1)
    Next
    
    For s = 0 To 3
        EditMap.Levels(Mode).MapData(Index).PressCheck(s) = False
        EditMap.Levels(Mode).MapData(Index).PressType(0) = 0
        EditMap.Levels(Mode).MapData(Index).PressTime(s) = 0
    Next
    
    With EditMap.Levels(Mode).MapData(Index)
        .PressCheck(Line) = True
        .PressType(0) = IIf(Late > 200, 1, 0)
        .StartTime = PressDown
        If Late > 200 Then .PressTime(Line) = Pos
    End With
 
    DrawHitPic Index
    
    PlayingLevel = EditMap.Levels(Mode)
End Sub
Sub EraseHitPic(Index As Integer)
    i = Index
    For s = 0 To 3
        If EditMap.Levels(Mode).MapData(i).PressCheck(s) Then
            If EditMap.Levels(Mode).MapData(i).PressType(0) = 0 Then
                BackImg.ImageByIndex(1).PresentWithClip HitPic.CompatibleDC, EditMap.Levels(Mode).MapData(i).StartTime / 1000 / GameBGM.max * 620, s * 10, 0, 0, 1, 10
            Else
                BackImg.ImageByIndex(1).PresentWithClip HitPic.CompatibleDC, EditMap.Levels(Mode).MapData(i).StartTime / 1000 / GameBGM.max * 620, s * 10, 0, 0, (EditMap.Levels(Mode).MapData(i).PressTime(s) - EditMap.Levels(Mode).MapData(i).StartTime) / 1000 / GameBGM.max * 620, 10
            End If
        End If
    Next
End Sub
Sub DrawHitPic(Index As Integer)
    i = Index
    For s = 0 To 3
        If EditMap.Levels(Mode).MapData(i).PressCheck(s) Then
            If EditMap.Levels(Mode).MapData(i).PressType(0) = 0 Then
                BackImg.ImageByIndex(1).PresentWithClip HitPic.CompatibleDC, EditMap.Levels(Mode).MapData(i).StartTime / 1000 / GameBGM.max * 620, 0, 0, 0, 1, 40
                BackImg.ImageByIndex(IIf(EditMap.Levels(Mode).MapData(i).PressType(0) = 0, 8, 7)).PresentWithClip HitPic.CompatibleDC, EditMap.Levels(Mode).MapData(i).StartTime / 1000 / GameBGM.max * 620, s * 10, 0, 0, 1, 10
            Else
                BackImg.ImageByIndex(1).PresentWithClip HitPic.CompatibleDC, EditMap.Levels(Mode).MapData(i).StartTime / 1000 / GameBGM.max * 620, 0, 0, 0, 1, 40
                BackImg.ImageByIndex(IIf(EditMap.Levels(Mode).MapData(i).PressType(0) = 0, 8, 7)).PresentWithClip HitPic.CompatibleDC, EditMap.Levels(Mode).MapData(i).StartTime / 1000 / GameBGM.max * 620, s * 10, 0, 0, (EditMap.Levels(Mode).MapData(i).PressTime(s) - EditMap.Levels(Mode).MapData(i).StartTime) / 1000 / GameBGM.max * 620, 10
            End If
        End If
    Next
End Sub
Sub ReDrawHitPic()
    BackImg.ImageByIndex(1).PresentWithClip HitPic.CompatibleDC, 0, 0, 0, 0, 620, 40

    For i = 1 To UBound(EditMap.Levels(Mode).MapData)
        For s = 0 To 3
            If EditMap.Levels(Mode).MapData(i).PressCheck(s) Then
                If EditMap.Levels(Mode).MapData(i).PressType(0) = 0 Then
                    BackImg.ImageByIndex(IIf(EditMap.Levels(Mode).MapData(i).PressType(0) = 0, 8, 7)).PresentWithClip HitPic.CompatibleDC, EditMap.Levels(Mode).MapData(i).StartTime / 1000 / GameBGM.max * 620, s * 10, 0, 0, 1, 10
                Else
                    BackImg.ImageByIndex(IIf(EditMap.Levels(Mode).MapData(i).PressType(0) = 0, 8, 7)).PresentWithClip HitPic.CompatibleDC, EditMap.Levels(Mode).MapData(i).StartTime / 1000 / GameBGM.max * 620, s * 10, 0, 0, (EditMap.Levels(Mode).MapData(i).PressTime(s) - EditMap.Levels(Mode).MapData(i).StartTime) / 1000 / GameBGM.max * 620, 10
                End If
            End If
        Next
    Next
End Sub
Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    HitPic.Create TargetDC, 620, 40
    GdipGraphicsClear HitPic.Graphics, argb(255, 255, 255, 255)
End Sub


