VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DownPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim ScrollY As Long
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "DownPage")
End Property

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "DownPage"
    
    Call DrawPage0
    
    Draw.Present GameDraw.CompatibleDC, 0, 0
End Sub
Public Sub DrawPage0()
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, HotAlpha
    
    Dim DrawY As Long, temp() As String, MapL() As String, MaxW As Long
    MapL = Split(MapList, "*")
    
    MaxW = (UBound(SongList) * 50)
    If MaxW > GWH Then MaxW = MaxW - GWH
    DrawY = 40 - ScrollY / 100 * MaxW
    'MainImg.Image("downloader.png").Present Draw.CompatibleDC, 40, DrawY
    DrawY = DrawY + 80
    If DrawY >= -25 And DrawY <= GWH Then GameFont.DrawText Draw.Graphics, 55, DrawY, 400, 25, "点击下载你喜欢的谱子吧！", argb(255, 255, 255, 255), near, 18, Bold
    
    DrawY = DrawY + 25
    If DrawY >= -25 And DrawY <= GWH Then GameFont.DrawText Draw.Graphics, 55, DrawY, GWW - 50, 25, "* 在编辑器中上传你的谱子！", argb(255, 232, 232, 232), near, 16
    DrawY = DrawY + 25
    
    If MapList = "" Then
        DrawY = DrawY + 25
        GameFont.DrawText Draw.Graphics, 55, DrawY, GWW - 120, 25, "正在加载谱子列表...", argb(255, 255, 255, 255), near, 16
    Else
        For i = UBound(MapL) To 1 Step -1
            If MapL(i) <> "" Then
                temp = Split(MapL(i), ">")
                DrawY = DrawY + 25
                If DrawY >= -25 And DrawY <= GWH Then
                    SetClickArea2 55, DrawY, GWW - 50, 24
                    If IsMouseIn Then BackImg.ImageByIndex(1).PresentWithClip Draw.CompatibleDC, 55, DrawY, 0, 0, GWW - 120, 25, 120
                    GameFont.DrawText Draw.Graphics, 55, DrawY, GWW - 120, 25, temp(0) & " - " & temp(1), IIf(IsMouseIn, argb(255, 255, 255, 255), argb(180, 255, 255, 255)), near, 16
                    If IsClick = True Then
                        ShellExecuteA MainWindow.hwnd, "open", temp(2), "", "", SW_SHOW
                    End If
                End If
            End If
            If DrawY > GWH Then Exit For
        Next
    End If
    
    CtrlImg.ImageByIndex(BMCtrl.BMVScroll).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 860, 100, ScrollY, 0, 16, Regular, VScrollBar2, , 100

End Sub

Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
End Sub



