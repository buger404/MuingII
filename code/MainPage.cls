VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim OpenTime As Long
Dim MusicR(10) As MusicData '11 FPS够了
Dim MusicR2(3) As MusicData
Dim MusicT As Long, MusicT2 As Long
Dim MusicP As Long, MusicP2 As Long
Dim MusicT3 As Long
Dim Played As Boolean
Dim ReNameMode As Boolean
Private Type MusicData
    Datas(60) As Long
    Time As Long
    Plus As Long
    Lower As Boolean
    Circles As Boolean
End Type
Dim Brush3 As Long, Brush As Long, Path As Long
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "MainPage")
End Property
Sub FileDrop(List As String)
    'On Error GoTo sth

    Dim data() As String
    Dim FSO As Object, tempMap As MuingII_MapFile, Path As String
    Dim ESongList As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    data = Split(List, vbCrLf)
    
    For i = 0 To UBound(data) - 1
        Path = data(i)
    
        If Dir(App.Path & "\temp\", vbDirectory) <> "" Then FSO.DeleteFolder App.Path & "\temp"
        
        Export Path, App.Path & "\temp\"
        
        Open App.Path & "\temp\map.mu" For Binary As #1
        Get #1, , tempMap
        Close #1
        
        Export Path, SongPath & "\" & tempMap.MapID & " - " & tempMap.Title & " - " & tempMap.Maker
        ESongList = ESongList & tempMap.MapID & " - " & tempMap.Title & " - " & tempMap.Maker & " - " & tempMap.Artist & vbCrLf
        
        If Dir(App.Path & "\temp\", vbDirectory) <> "" Then FSO.DeleteFolder App.Path & "\temp"
    Next
    
    GameNotify.Message MuingSuccess, "成功导入了 " & vbCrLf & ESongList, "好！"
    GetSongList
    If UBound(SongList) > 0 Then
        PlayRandomSong
    End If
    
sth:
    If Err.Number <> 0 Then
        Err.Clear
        GameNotify.Message MuingError, "谱子导入失败！", "QWQ"
        On Error Resume Next
        Close #1
    End If
End Sub

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "MainPage"
    
    If OpenTime = 0 Or GetTickCount - OpenTime <= 3000 Then
        If OpenTime = 0 Then OpenTime = GetTickCount
        Call DrawPage0
    Else
        If Played = False Then
            ShowTask = True: Played = True: GameCore.FadePage "MainPage"
            If Settings(MuingII_Settings.MuingNoServer) = 0 Then
                If DebugMode Then
                    Sock.Connect "127.0.0.1", 4049
                Else
                    Sock.Connect "free.idcfengye.com", 17028
                End If
            Else
                GiveupConnect = True
            End If
            
        End If
        Call DrawPage1
    End If
    Draw.Present GameDraw.CompatibleDC, 0, 0
End Sub

Public Sub DrawPage0()
    Dim Pro As Single
    Pro = (GetTickCount - OpenTime) / 3000
    Pro = cubicCurves(Pro, 0, 1, 1, 1)
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, Pro * 220
    
    GdipSetSolidFillColor Brush, argb(255, 9, 233, 179)
    GdipFillEllipse Draw.Graphics, Brush, GWW / 2 - (Pro * 150) / 2, GWH / 2 - (Pro * 150) / 2 + 27, (Pro * 150), (Pro * 150)
    GdipSetSolidFillColor Brush, argb(Pro * 255, 255, 255, 255)
    GdipFillEllipse Draw.Graphics, Brush, GWW / 2 - (Pro * 150) / 2, GWH / 2 - (Pro * 150) / 2 + 27, (Pro * 150), (Pro * 150)
    
    If NowSong <> 0 Then SongList(NowSong).SongPicCircle.Present Draw.CompatibleDC, GWW / 2 - 64, GWH / 2 - 64 + 27, (Pro * 120)
    
    Dim MusicD() As Single
    MusicD = GameBGM.GetMusicBar
    
    Dim Hot As Long, HotW As Long
    HotW = Int(GWW / 30)
    
    If GameBGM.PlayState = Playing Then
        For i = 0 To 30
            Hot = MusicD(i) / 8 * Pro
            BackImg.ImageByIndex(1).PresentWithClip Draw.CompatibleDC, i * HotW, GWH - Hot, 0, 0, HotW, Hot, Hot / (GWH / 2) * 255
        Next
    End If
    
    BackImg.ImageByIndex(2).Present Draw.CompatibleDC, 0, 0, 255 - 255 * Pro
End Sub
Public Sub DrawPage1()
    Dim MusicD() As Single, s As Integer, OranPlus As Single, NowHot As Long
    MusicD = GameBGM.GetMusicBar
    If GetTickCount - MusicP2 >= 64 Then
        If GameBGM.PlayState = Playing Then
            If GetTickCount - MusicR(MusicP).Time < 600 Then Debug.Print Now, "ffff"
            MusicR(MusicP).Time = GetTickCount
            OranPlus = MusicR(MusicP).Plus: MusicR(MusicP).Plus = 0
            For i = 0 To 60
                If MusicT3 > 360 Then MusicT3 = MusicT3 - 360
                s = i + MusicT3 / 360 * 60
                If s > 60 Then s = s - 60
                MusicR(MusicP).Datas(s) = MusicD(i / 60 * UBound(MusicD))
                MusicR(MusicP).Plus = MusicR(MusicP).Plus + MusicR(MusicP).Datas(s)
            Next
            MusicR(MusicP).Circles = (OranPlus * 2 - MusicR(MusicP).Plus < 0)
            MusicR(MusicP).Lower = (OranPlus * 2 - MusicR(MusicP).Plus > 800)
            If GetTickCount - MusicT >= 250 Then
                Dim last As Long
                last = MusicR2(MusicT2).Plus
                MusicT2 = MusicT2 + 1
                If MusicT2 > 3 Then MusicT2 = 3
                MusicR2(MusicT2) = MusicR(MusicP)
                If Abs(MusicR2(MusicT2).Plus - last) > 400 Then
                    MusicT3 = MusicT3 + 0
                End If
                If MusicT3 > 360 Then MusicT3 = MusicT3 - 360
            End If
        End If
        MusicP = MusicP + 1
        If MusicP > UBound(MusicR) Then MusicP = 0
        MusicP2 = GetTickCount
    End If
    
    If GameBGM.PlayState = Playing Then
        For i = 0 To 60
            NowHot = NowHot + MusicD(i / 60 * UBound(MusicD))
        Next
    End If
    
    MusicT3 = MusicT3 + (NowHot / 3000)
    
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, 210 - (NowHot / 300)
    
    Dim x As Long, r As Single
    Dim RX1 As Long, RX2 As Long, RA As Long, RD As Single, RD3 As Long
    Dim LowC As Long, HighC As Long
    'GdipCreateSolidFill 0, Brush3
    
    RX1 = -258 + 200: RX2 = 55 + 200
    
    '170 r (170 / 2) = 85
    For i = 0 To UBound(MusicR)
        RD = (GetTickCount - MusicR(i).Time) / 700
        If RD < 1 Then
            'GdipSetSolidFillColor Brush3, argb(255 - 255 * RD, 0, 234, 118)
            Dim PW As Long
            PW = (NowHot / 250) * 2
            
            If MusicR(i).Circles = True Then
                HighC = HighC + 1
                RX1 = 150 + 400 * RD
                GdipSetSolidFillColor Brush, argb(60 - 60 * RD, 255, 255, 255)
                'GdipDrawEllipse Draw.Graphics, Pen, -263 + (800 - RX1) / 2, 58 + (800 - RX1) / 2, RX1, RX1
                GdipFillEllipse Draw.Graphics, Brush, GWW / 2 - 150 / 2 + (150 - RX1) / 2, GWH / 2 - 150 / 2 + (150 - RX1) / 2 + 27, RX1, RX1
            End If
            
            If MusicR(i).Lower = True Then LowC = LowC + 1
            
            For s = 0 To 60
                If s < 30 Then
                    RA = -180 + MusicT3 + s / 30 * 180
                Else
                    RA = 0 + MusicT3 + (s - 30) / 30 * 180
                End If
                RX1 = PW + 300 * (MusicR(i).Datas(s) / 600)
                RX1 = RX1 * (1 - RD) + 150
                If RX1 > 150 Then
                    RX2 = (RX1 - 150) / 2
                    RD3 = (400 - RX1) / 2
                    GdipAddPathArc Path, GWW / 2 - 400 / 2 + RD3, GWH / 2 - 400 / 2 + RD3 + 27, RX1, RX1, RA, 3
                    GdipAddPathArc Path, GWW / 2 - 400 / 2 + RD3 + RX2, GWH / 2 - 400 / 2 + RD3 + RX2 + 27, 150, 150, RA + 3, -3
                    GdipFillPath Draw.Graphics, Brush3, Path
                    GdipResetPath Path
                End If
            Next
        
        End If
    Next
    
    GdipSetSolidFillColor Brush, argb(255, 9, 233, 179)
    GdipFillEllipse Draw.Graphics, Brush, GWW / 2 - 150 / 2 - PW / 2, GWH / 2 - 150 / 2 + 27 - PW / 2, 150 + PW, 150 + PW
    GdipSetSolidFillColor Brush, argb(NowHot / 3000 * 255, 255, 255, 255)
    GdipFillEllipse Draw.Graphics, Brush, GWW / 2 - 150 / 2 - PW / 2, GWH / 2 - 150 / 2 + 27 - PW / 2, 150 + PW, 150 + PW
    
    If NowSong <> 0 Then SongList(NowSong).SongPicCircle.Present Draw.CompatibleDC, GWW / 2 - 64, GWH / 2 - 64 + 27, 55 + (NowHot / 10000) * 100
    
    GameFont.DrawText Draw.Graphics, 0, GWH - 70, GWW - 10, 30, "网络延迟 " & CLate & " ms", argb(100, 255, 255, 255), far, 14
    GameFont.DrawText Draw.Graphics, 0, GWH - 50, GWW - 10, 30, "Made by Error 404 (QQ 1361778219)", argb(100, 255, 255, 255), far, 14
    GameFont.DrawText Draw.Graphics, 0, GWH - 30, GWW - 10, 30, "Muing II beta 19.0105 / Wonder Version", argb(150, 255, 255, 255), far, 14, Bold
    
    Dim Hot As Long, HotW As Long
    HotW = Int(GWW / 30)
    
    If GameBGM.PlayState = Playing Then
        For i = 0 To 30
            Hot = MusicD(i) / 8
            BackImg.ImageByIndex(1).PresentWithClip Draw.CompatibleDC, i * HotW, GWH - Hot, 0, 0, HotW, Hot, Hot / (GWH / 2) * 255
        Next
    End If
    
    If Connected = False And GiveupConnect = False Then
        Select Case Sock.State
            Case 1
            GameNotify.TipText = "客户端不动了。"
            Case 2
            GameNotify.TipText = "准备连接到服务器..."
            Case 3
            GameNotify.TipText = "服务器不动了。"
            Case 4
            GameNotify.TipText = "正在解析服务器域名..."
            Case 5
            GameNotify.TipText = "成功识别服务器主机..."
            Case 6
            GameNotify.TipText = "正在连接到服务器..."
            Case 7
            GameNotify.TipText = "已建立连接。"
            Case 8
            GameNotify.TipText = "服务器正在关闭..."
            Case 9
            GameNotify.TipText = "连接错误。"
        End Select
    End If
End Sub

Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    AddMember Me
    GdipCreateLineBrush NewPointF(0, 0), NewPointF(GWW, GWH), argb(55, 0, 174, 245), argb(55, 0, 254, 138), WrapModeTile, Brush3
    GdipCreateSolidFill 0, Brush
    GdipCreatePath FillModeWinding, Path
End Sub
Sub Dispose()
    GdipDeleteBrush Brush3: GdipDeleteBrush Brush
    GdipDeletePath Path
End Sub
