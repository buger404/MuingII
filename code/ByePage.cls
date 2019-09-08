VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ByePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim OpenTime As Long
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "ByePage")
End Property

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "ByePage"
    
    Dim Pro As Single, Brush As Long
    Pro = (GetTickCount - OpenTime) / 2000
    Pro = cubicCurves(Pro, 0, 0, 0, 1)
    Pro = 1 - Pro
    If Pro < 0 Then Pro = 0
    If Pro > 1 Then Pro = 1
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, Pro * 220
    
    GdipCreateSolidFill 0, Brush
    GdipSetSolidFillColor Brush, argb(255, 9, 233, 179)
    GdipFillEllipse Draw.Graphics, Brush, GWW / 2 - (Pro * 150) / 2, GWH / 2 - (Pro * 150) / 2 + 27, (Pro * 150), (Pro * 150)
    GdipSetSolidFillColor Brush, argb(Pro * 255, 255, 255, 255)
    GdipFillEllipse Draw.Graphics, Brush, GWW / 2 - (Pro * 150) / 2, GWH / 2 - (Pro * 150) / 2 + 27, (Pro * 150), (Pro * 150)
    GdipDeleteBrush Brush
    
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
    
    If year(Now) >= 2019 And year(Now) <= 2022 Then
        If year(Now) = 2019 Then
            If Month(Now) >= 8 Then BackImg.ImageByIndex(5).Present Draw.CompatibleDC, 0, 0, 255 - 255 * Pro
        ElseIf year(Now) = 2022 Then
            If Month(Now) < 6 Then BackImg.ImageByIndex(5).Present Draw.CompatibleDC, 0, 0, 255 - 255 * Pro
        Else
            BackImg.ImageByIndex(5).Present Draw.CompatibleDC, 0, 0, 255 - 255 * Pro
        End If
    Else
        BackImg.ImageByIndex(4).Present Draw.CompatibleDC, 0, 0, 255 - 255 * Pro
    End If
    
    If OpenTime = 0 Then OpenTime = GetTickCount
    If GetTickCount - OpenTime >= 2000 Then CloseTime = True
    GameBGM.volume = 0.5 - (GetTickCount - OpenTime) / 2000 * 0.5
    
    Draw.Present GameDraw.CompatibleDC, 0, 0
End Sub

Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
End Sub




