Attribute VB_Name = "DrawingLevel"
Dim LDraw As Images, LDraw2 As Images
Public LastHit As Integer, ActiveHit  As Integer, EditIndex As Integer
Dim EffectTime(3) As Long, EffectColor(3) As Integer
Public Score As Double, Combo As Long, NowCombo As Long
Public Grade As String
Public ExcellentCount As Long, GoodCount As Long, OKCount As Long, MissCount As Long
Public PlayTime As Long
Public HitLate As Long, ObjImg(4) As Images, ObjImg2(4) As Images
Public HP As Single
Public EffectTipTime As Long, EffectTipIndex As Integer, EffectTipImg(3) As Images
Public PlayingLevel As MuingII_Map
Public Accuracy As Single
Public HitDown As Long

Public PressLock As Integer, LongPress As Boolean
Public SpacePress As Boolean

Public HitDownLate As Single

Dim AutoHit As Long
Sub LevelDrawing(DC As Long, Graphics As Long, X As Long, Y As Long, EditMode As Boolean)
    'If HP < 0 Then mNowShow = "SoloPage": Exit Sub

    'If EditMode = True Then PlayingLevel = EditMap.Levels(EditMode)
    
    Dim EffectLig As Long
    
    If (GetTickCount - EffectTipTime) <= 400 Then
        EffectLig = 50 - (GetTickCount - EffectTipTime) / 400 * 50
    End If
    
    'NowHot = 0
    BackImg.ImageByIndex(3).Present DC, 0, 0, HotAlpha - EffectLig
    
    If LDraw Is Nothing Then
        Set LDraw = New Images: LDraw.Create TargetDC, GWW, GWH
        Set LDraw2 = New Images: LDraw2.Create TargetDC, GWW, GWH
    End If
    
    Dim MoveTime As Single, MoveTime2 As Single, Late As Single, Late2 As Single
    Dim Active As Integer, ActiveLine(3) As Boolean, ShowLine(3) As Boolean
    Dim DrawX As Long, DrawX2 As Long
    Dim Line As Integer, i As Integer
    
    Active = -1
    
    GdipGraphicsClear LDraw.Graphics, 0
    GdipGraphicsClear LDraw2.Graphics, 0
    
    MoveTime = GWW / (250 * PlayingLevel.ObjSpeed)
    MoveTime2 = 47 * 2 / (250 * PlayingLevel.ObjSpeed)
    
    Dim Alpha As Long, BGMPos As Single, PrintCount As Long, s As Integer
    BGMPos = GameBGM.Pos
    
    If HP > 99 Then
        If UBound(PlayingLevel.MapData) >= 1 Then
            If (PlayingLevel.MapData(1).StartTime / 1000 + MoveTime2) - BGMPos > 3 Then
                With PlayImg.Image("skip.png")
                    .Present LDraw.CompatibleDC, 770, 320
                    .SetClickArea 770 + X, 320 + Y
                    If IsClick Or SpacePress Then
                        MenuSnd2.Play: GameBGM.Pos = (PlayingLevel.MapData(1).StartTime / 1000 + MoveTime2) - 2.9
                    End If
                End With
                SpacePress = False
            End If
        End If
    End If


    For i = LastHit + 1 To UBound(PlayingLevel.MapData)
        Late = BGMPos - (PlayingLevel.MapData(i).StartTime / 1000 + MoveTime2)
        DrawX = -(Late / MoveTime * (GWW + 47 * 2)) - 61
        DrawX2 = 0
        If DrawX >= -45 And DrawX <= GWW Then
            If PlayingLevel.MapData(i).PressType(0) = 1 Then
                For s = 0 To 3
                    If PlayingLevel.MapData(i).PressCheck(s) Then
                        Late2 = BGMPos - (PlayingLevel.MapData(i).PressTime(s) / 1000 + MoveTime2)
                        DrawX2 = -(Late2 / MoveTime * (GWW + 47 * 2)) - 61
                        Line = s: Exit For
                    End If
                Next
                If PressLock = Line And i = LastHit + 1 And EditMode = False Then
                    If GetTickCount - EffectTipTime >= 200 Then
                        EffectTipTime = GetTickCount
                        SetEffect Line, Int(ActiveHit / 10) Mod 4
                    End If
                    PlayingLevel.MapData(i).StartTime = GameBGM.Pos * 1000
                    If PlayingLevel.MapData(i).StartTime >= PlayingLevel.MapData(i).PressTime(Line) And (ModPower(0) Or EditMode) Then
                        HitObj Line
                        PressLock = -1
                    End If
                End If
            End If
            If ModPower(3) Then DrawX = cubicCurves(DrawX / GWW, 0, 1, 1, 1) * GWW
            If ModPower(3) And DrawX2 <> 0 Then DrawX2 = cubicCurves(DrawX2 / GWW, 0, 1, 1, 1) * GWW
            PrintCount = PrintCount + 1
            For s = 0 To 3
                If PlayingLevel.MapData(i).PressCheck(s) Then
                    If PlayingLevel.MapData(i).PressType(0) = 1 Then
                        With ObjImg2(Int(i / 10) Mod 4)
                            .PresentWithClip LDraw.CompatibleDC, DrawX + ObjImg(0).Width / 2, s * 100 + 41, 0, 0, DrawX2 - DrawX, .Height, IIf(i = EditIndex, 255, 100)
                        End With
                    End If
                    If EditMode = True Then
                        Alpha = IIf(i = EditIndex, 255, 100)
                        With ObjImg(Int(i / 10) Mod 4)
                            .Present LDraw.CompatibleDC, DrawX, s * 100 + 16, Alpha
                            If PlayingLevel.MapData(i).PressType(0) = 1 Then
                                .Present LDraw.CompatibleDC, DrawX2, s * 100 + 16, Alpha
                                SetClickArea2 DrawX, s * 100 + 20 + Y, DrawX2 - DrawX + .Width * 2, .Height
                            Else
                                .SetClickArea DrawX, s * 100 + 20 + Y
                            End If
                            If IsClick = True Then
                                EditIndex = i
                            End If
                        End With
                    Else
                        If ModPower(2) Then
                            If PlayingLevel.MapData(i).PressType(0) = 1 Then
                                With ObjImg2(Int(i / 10) Mod 4)
                                    .PresentWithClip LDraw.CompatibleDC, DrawX + ObjImg(0).Width / 2, s * 100 + 41, 0, 0, DrawX2 - DrawX, .Height, IIf(DrawX > 400, (DrawX - 400) / GWW * 255, 0)
                                End With
                            End If
                            ObjImg(Int(i / 10) Mod 4).Present LDraw.CompatibleDC, DrawX, s * 100 + 16, IIf(DrawX > 400, (DrawX - 400) / GWW * 255, 0)
                            If PlayingLevel.MapData(i).PressType(0) = 1 Then ObjImg(Int(i / 10) Mod 4).Present LDraw.CompatibleDC, DrawX2, s * 100 + 16, IIf(DrawX2 > 400, (DrawX2 - 400) / GWW * 255, 0)
                        Else
                            If PlayingLevel.MapData(i).PressType(0) = 1 Then
                                With ObjImg2(Int(i / 10) Mod 4)
                                    .PresentWithClip LDraw.CompatibleDC, DrawX + ObjImg(0).Width / 2, s * 100 + 41, 0, 0, DrawX2 - DrawX, .Height, 255
                                End With
                            End If
                            ObjImg(Int(i / 10) Mod 4).Present LDraw.CompatibleDC, DrawX, s * 100 + 16, 255
                            If PlayingLevel.MapData(i).PressType(0) = 1 Then ObjImg(Int(i / 10) Mod 4).Present LDraw.CompatibleDC, DrawX2, s * 100 + 16, 255
                        End If
                    End If
                    
                    If Active = -1 Or Active = i Then
                        ActiveHit = i
                        ActiveLine(s) = True: Active = i
                        If ModPower(2) Then
                            ObjImg(4).Present LDraw.CompatibleDC, DrawX - 1, s * 100 + 16, IIf(DrawX > 400, (DrawX - 400) / GWW * 255, 0)
                            If PlayingLevel.MapData(i).PressType(0) = 1 Then ObjImg(4).Present LDraw.CompatibleDC, DrawX2 - 1, s * 100 + 16, IIf(DrawX2 > 400, (DrawX2 - 400) / GWW * 255, 0)
                        Else
                            ObjImg(4).Present LDraw.CompatibleDC, DrawX - 1, s * 100 + 16
                            If PlayingLevel.MapData(i).PressType(0) = 1 Then ObjImg(4).Present LDraw.CompatibleDC, DrawX2 - 1, s * 100 + 16
                        End If
                    End If
                    ShowLine(s) = True
                    If EditMode = True Or ModPower(0) Then
                        Late = (PlayingLevel.MapData(i).StartTime / 1000) - BGMPos
                        If Abs(Late) <= 0.03 And LastHit < i Then
                            If PlayingLevel.MapData(i).PressType(0) = 1 Then
                                If PressLock = -1 Then PressLock = Line: HitDownLate = Abs((PlayingLevel.MapData(ActiveHit).StartTime / 1000) - GameBGM.Pos): HitObj2 s
                            Else
                                HitDownLate = Abs((PlayingLevel.MapData(ActiveHit).StartTime / 1000) - GameBGM.Pos)
                                HitObj s
                                HitLate = (HitLate + (BGMPos * 1000 - PlayingLevel.MapData(i).StartTime)) / 2
                            End If
                            If ModPower(0) Then AutoHit = GetTickCount
                        End If
                    End If
                End If
            Next
        ElseIf DrawX > GWW Then
            Exit For
        ElseIf DrawX < 0 Then
            If EditMode = False Then
                If Abs((PlayingLevel.MapData(i).StartTime / 1000) - GameBGM.Pos) > 0.4 Then
                    LastHit = i
                    Call OnMiss
                End If
            End If
        End If
    Next
    
    Alpha = IIf(EditMode, 100, 255)
    
    If LastHit = 0 Then
        Accuracy = 1
        Grade = "-"
    Else
        Accuracy = (ExcellentCount * 5 + GoodCount * 3 + OKCount * 1) / (LastHit * 5)
        Grade = GetGrade
    End If
    
    If Settings(MuingII_Settings.MuingShowAllLines) = 1 Then
        For i = 0 To 3
            ShowLine(i) = True
        Next
    End If
    
    If ShowLine(0) Then PlayImg.Image(IIf(ActiveLine(0), "focus", "play") & "frame1.png").Present LDraw2.CompatibleDC, 0, 0, Alpha
    If ShowLine(1) Then PlayImg.Image(IIf(ActiveLine(1), "focus", "play") & "frame2.png").Present LDraw2.CompatibleDC, 0, 100, Alpha
    If ShowLine(2) Then PlayImg.Image(IIf(ActiveLine(2), "focus", "play") & "frame3.png").Present LDraw2.CompatibleDC, 0, 200, Alpha
    If ShowLine(3) Then PlayImg.Image(IIf(ActiveLine(3), "focus", "play") & "frame4.png").Present LDraw2.CompatibleDC, 0, 300, Alpha
    
    If Not ModPower(0) Then
        For i = 0 To 3
            SetClickArea2 0 + X, 100 * i + 5 + Y, GWW, 100
            If Not EditMode Then
                If IsMouseDownNoKeep And PressLock = -1 Then
                    HitDownLate = Abs((PlayingLevel.MapData(ActiveHit).StartTime / 1000) - GameBGM.Pos)
                    HitObj2 i: PressLock = i
                End If
                If IsMouseUp Then
                    HitObj i: PressLock = -1
                End If
            End If
        Next
    End If
    
    For i = 0 To 3
        If (GetTickCount - EffectTime(i)) <= 400 Then
            PlayImg.Image("light" & EffectColor(i) & ".png").Present LDraw.CompatibleDC, 0, 100 * i + 5, 255 - (GetTickCount - EffectTime(i)) / 400 * 255
        End If
    Next
    
    If (GetTickCount - EffectTipTime) <= 400 Then
        With EffectTipImg(EffectTipIndex)
            'BackImg.ImageByIndex(1).PresentWithClip LDraw.CompatibleDC, 0, 300 / 2 - .Height / 2 - 7, 0, 0, GWW, .Height + 7, 80 - cubicCurves((GetTickCount - EffectTipTime) / 400, 0, 0, 0, 1) * 80
            .Present LDraw.CompatibleDC, GWW / 2 - .Width / 2, 300 / 2 - .Height / 2, 255 - (GetTickCount - EffectTipTime) / 400 * 255
        End With
    End If
    
    If ModPower(0) Then
        Dim Auto As Long
        Auto = GetTickCount - AutoHit
        If Auto > 400 Then Auto = 400
        With BMImg.ImageByIndex(1)
            .Present LDraw.CompatibleDC, GWW - .Width + 61, 348, Auto / 400 * 255
        End With
        With BMImg.ImageByIndex(2)
            .Present LDraw.CompatibleDC, GWW - .Width + 61, 348, 255 - Auto / 400 * 255
        End With
        GameFont.DrawText LDraw2.Graphics, 0, 398, GWW, 30, "正在观看黑嘴玩耍" & SongList(NowSong).Info.Title & "...", argb(255, 255, 255, 255), center, 16
    End If
    
    LDraw.Present LDraw2.CompatibleDC, 0, 0
    
    If EditMode = True Then
        GameFont.DrawText LDraw2.Graphics, 30, 0, 200, 30, "即时物件数量：" & PrintCount, argb(255, 255, 255, 255), near, 16
        GameFont.DrawText LDraw2.Graphics, 30, 25, 200, 30, "打击延迟：" & HitLate & "ms", argb(255, 255, 255, 255), near, 16
        'CtrlImg.ImageByIndex(1).PresentWithCtrl DC, Graphics, 30, 530, "1x", argb(255, 255, 255, 255), 14, Regular, Button
        'If IsClick Then GameBGM.SetPlayRate 1
        'CtrlImg.ImageByIndex(1).PresentWithCtrl DC, Graphics, 30 + 150, 530, "1.5x", argb(255, 255, 255, 255), 14, Regular, Button
        'If IsClick Then GameBGM.SetPlayRate 1.5
        'CtrlImg.ImageByIndex(1).PresentWithCtrl DC, Graphics, 30 + 150 * 2, 530, "2x", argb(255, 255, 255, 255), 14, Regular, Button
        'If IsClick Then GameBGM.SetPlayRate 2
    End If
    
    If LastHit + 1 <= UBound(PlayingLevel.MapData) And GameBGM.PlayState = Playing Then
        Dim Wait As Single, MaxWait As Single, W As Long
        Wait = (PlayingLevel.MapData(LastHit + 1).StartTime / 1000) - BGMPos

        If Wait >= 3 Then
            MaxWait = (PlayingLevel.MapData(LastHit + 1).StartTime / 1000) - IIf(LastHit = 0, 0, (PlayingLevel.MapData(LastHit).StartTime / 1000))
            If MaxWait > 3 Then
                GameFont.DrawText LDraw2.Graphics, 0, 90, GWW, 50, "Relax", argb(255, 255, 255, 255), center, 24, Bold
                With PlayImg.Image("relax.png")
                    W = .Width * ((Wait - 3) / (MaxWait - 3))
                    .PresentWithClip LDraw2.CompatibleDC, GWW / 2 - W / 2, 140, 0, 0, W, .Height
                End With
                GameFont.DrawText LDraw2.Graphics, 0, 170, GWW, 60, Int(Wait - 3) + 1 & "s", argb(255, 255, 255, 255), center, 24, Bold
            End If
        End If
    End If
    
    HP = HP + 0.08
    If HP > 100 Then HP = 100
    
    LDraw2.Present DC, X, Y
End Sub
Sub HitObj(Line As Integer)
    If ActiveHit = LastHit Then Exit Sub
    If ActiveHit > UBound(PlayingLevel.MapData) Then Exit Sub
    
    'If LongPress = True Then LongPress = False: Exit Sub
    
    Dim Hited As Boolean, Late As Single, DrawX As Long, BasicScore As Long
    Dim TrueLine As Integer, AddScore As Boolean
    Dim MoveTime As Single, MoveTime2 As Single
    AddScore = (PlayingLevel.MapData(ActiveHit).PressType(0) = 0) Or (PlayingLevel.MapData(ActiveHit).PressTime(Line) = PlayingLevel.MapData(ActiveHit).StartTime)
    If PlayingLevel.MapData(ActiveHit).PressType(0) = 1 Then
        If PlayingLevel.MapData(ActiveHit).PressTime(Line) - PlayingLevel.MapData(ActiveHit).StartTime <= 400 Then
            PlayingLevel.MapData(ActiveHit).StartTime = PlayingLevel.MapData(ActiveHit).PressTime(Line)
            AddScore = True
        End If
    End If
    MoveTime = GWW / (250 * PlayingLevel.ObjSpeed)
    MoveTime2 = 47 * 2 / (250 * PlayingLevel.ObjSpeed)
    Late = GameBGM.Pos - (PlayingLevel.MapData(ActiveHit).StartTime / 1000 + MoveTime2)
    DrawX = -(Late / MoveTime * (GWW + 47 * 2)) - 61
    
    Late = HitDownLate
    
    For i = 0 To 3
        If PlayingLevel.MapData(ActiveHit).PressCheck(i) Then TrueLine = i
    Next
    
    If DrawX >= GWW / 2 Then  '玩家太急了但是应该是手误
        '放过玩家吧
        Exit Sub
    End If
    
    If Late >= 0 And Late <= 0.1 Then 'Excellent判定区域
        EffectTipIndex = 0: BasicScore = 300: Hited = True
        If AddScore Then ExcellentCount = ExcellentCount + 1
    End If
    If Late >= 0.1 And Late <= 0.2 Then 'Great判定区域
        EffectTipIndex = 1: BasicScore = 150: Hited = True
        If AddScore Then GoodCount = GoodCount + 1
    End If
    If Late >= 0.2 And Late <= 0.4 Then 'Okay判定区域
        EffectTipIndex = 2: BasicScore = 50: Hited = True
        If AddScore Then OKCount = OKCount + 1
    End If
    
    Hited = Hited And (TrueLine = Line)
    
    If Hited = True Then
        EffectTipTime = GetTickCount
        If AddScore Then
            NowCombo = NowCombo + 1
            Score = Score + BasicScore * (NowCombo / 8): PressLock = -1: LastHit = ActiveHit
            If NowCombo > Combo Then Combo = NowCombo
        End If
        
        If PlayingLevel.MapData(ActiveHit).PressType(0) = 0 Then
            If ModPower(0) And Settings(MuingII_Settings.MuingUseDog) = 0 Then
                AutoSnd.Play
            Else
                If (GetTickCount - HitDown) <= 110 Then
                    HitSnd.Play
                Else
                    HitSnd2.Play
                End If
            End If
        End If
        SetEffect Line, Int(ActiveHit / 10) Mod 4
    Else
        If AddScore Then PressLock = -1: LastHit = ActiveHit
        Call OnMiss
    End If
    
    'Debug.Print Now, LastHit, PressLock
End Sub
Sub HitObj2(Line As Integer)
    If PlayingLevel.MapData(ActiveHit).PressType(0) = 1 Then
        Dim Late As Single
        
        Late = HitDownLate
        
        If Late >= 0 And Late <= 0.1 Then 'Excellent判定区域
           EffectTipIndex = 0
        End If
        If Late >= 0.1 And Late <= 0.2 Then 'Great判定区域
           EffectTipIndex = 1
        End If
        If Late >= 0.2 And Late <= 0.4 Then 'Okay判定区域
           EffectTipIndex = 2
        End If
         
        HitSnd2.Play
    End If
End Sub
Function GetGrade() As String
    Dim EP As Single, GP As Single, OP As Single, MP As Single
    EP = (ExcellentCount * 2 + GoodCount * 1.5 + OKCount) / (LastHit * 2)
    GP = GoodCount / LastHit
    OP = OKCount / LastHit
    MP = MissCount / LastHit
    
    If EP = 1 Then
        GetGrade = "SS"
    ElseIf EP >= 0.95 And MP <= 0.02 Then
        GetGrade = "S"
    ElseIf EP >= 0.85 And MP <= 0.05 Then
        GetGrade = "A"
    ElseIf EP >= 0.75 And MP <= 0.07 Then
        GetGrade = "B"
    ElseIf EP >= 0.6 And MP <= 0.1 Then
        GetGrade = "C"
    Else
        GetGrade = "D"
    End If
End Function
Sub OnMiss()
    PressLock = -1
    MissSnd.Play
    MissCount = MissCount + 1
    EffectTipTime = GetTickCount
    EffectTipIndex = 3
    NowCombo = 0: HP = HP - 10
End Sub
Sub SetEffect(ByVal Line As Integer, ByVal Color As Integer)
    EffectTime(Line) = GetTickCount
    EffectColor(Line) = Color
End Sub
Sub DrawMod(DC As Long, Graphics As Long, X As Long, Y As Long)
    Dim TipText(3) As String
    TipText(0) = "Auto : 自动完成一张谱子（0x）"
    TipText(1) = "Quick : 1.5倍速进行游戏（1.2x）"
    TipText(2) = "Fade : 距离屏幕左边越近，物件透明度越高（1.2x）"
    TipText(3) = "Drop : 奇特的物件移动方式（1.1x）"
    For i = 0 To 3
        With ModImg.Image("mod" & i & IIf(ModPower(i), "active", "") & ".png")
            .Present DC, X + i * .Width, 89
            .SetClickArea X + i * .Width, Y
            If IsMouseIn Then GameNotify.TipText = TipText(i)
            If IsClick Then MenuSnd2.Play: ModPower(i) = Not ModPower(i)
        End With
    Next
End Sub
Sub DrawMod2(DC As Long, Graphics As Long, X As Long, Y As Long)
    Dim X2 As Long
    X2 = X
    
    For i = 0 To 3
        If ModPower(i) Then
            With ModImg.Image("mod" & i & ".png")
                X2 = X2 - .Width
                .Present DC, X2, Y
            End With
        End If
    Next
End Sub
Sub DrawMod3(DC As Long, Graphics As Long, X As Long, Y As Long)
    Dim X2 As Long
    X2 = X
    
    For i = 0 To 3
        If ModPower(i) Then
            With ModImg.Image("mod" & i & "active.png")
                .Present DC, X2, Y
                X2 = X2 + .Width + 10
            End With
        End If
    Next
End Sub
