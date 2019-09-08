VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Screens() As Object, StopWork As Boolean
Dim TransTime As Long, TransPage As String, TransEffect As Integer
Dim FadeDraw As New Images
Sub FadePage(Page As String)
    GameDraw.Present FadeDraw.CompatibleDC, 0, 0
    TransTime = GetTickCount
    TransPage = Page
    TransEffect = 1
End Sub
Sub ChangePage(Page As String)
    TransTime = GetTickCount
    TransPage = Page
    TransEffect = 0
End Sub
Sub AddScreen(NewScreen As Object)
    ReDim Preserve Screens(UBound(Screens) + 1)
    Set Screens(UBound(Screens)) = NewScreen
End Sub
Sub Display()
    If DebugMode = False Then On Error GoTo CatchError
    
    If TransPage = "ByePage" Then StopWork = False
    If StopWork = True Then Exit Sub
    
    CtrlX = -1: CtrlY = -1: CtrlW = 0: CtrlH = 0
    NowEdit = -1
    
    If BackChangeTime <> 0 Then
        If GetPastTick(BackChangeTime) <= 1000 Then
            Dim Pro As Single
            Pro = cubicCurves(GetPastTick(BackChangeTime) / 1000, 0, 1, 1, 1)
            MainBack2.Present MainBack.CompatibleDC, 0, 0
            MainBack3.Present MainBack.CompatibleDC, 0, 0, Pro * 255
        Else
            MainBack3.Present MainBack.CompatibleDC, 0, 0
            BackChangeTime = 0
        End If
    End If
    
    Select Case TransEffect
        Case 0
            If GetTickCount - TransTime < 1000 Then
                UnClicked = True
                If GetTickCount - TransTime >= 500 Then
                   If mNowShow <> TransPage Then mNowShow = TransPage
                End If
            Else
                UnClicked = False
            End If
        Case 1
            If GetTickCount - TransTime < 400 Then
                If mNowShow <> TransPage Then mNowShow = TransPage
                UnClicked = True
            Else
                UnClicked = False
            End If
    End Select
    
    For i = 1 To UBound(Screens)
        If Screens(i).Visible Then Screens(i).Display
    Next

    Select Case TransEffect
        Case 0
            If GetTickCount - TransTime < 1000 Then
                If GetTickCount - TransTime < 500 Then
                    BackImg.ImageByIndex(1).Present GameDraw.CompatibleDC, 0, 0, cubicCurves((GetTickCount - TransTime) / 500, 0, 1, 1, 1) * 255
                ElseIf GetTickCount - TransTime >= 500 Then
                    BackImg.ImageByIndex(1).Present GameDraw.CompatibleDC, 0, 0, 255 - cubicCurves((GetTickCount - TransTime - 500) / 500, 0, 1, 1, 1) * 255
                End If
            End If
        Case 1
            If GetTickCount - TransTime < 400 Then FadeDraw.Present GameDraw.CompatibleDC, 0, 0, 255 - cubicCurves((GetTickCount - TransTime) / 400, 0, 1, 1, 1) * 255
    End Select
    
    'GameFont.DrawText GameDraw.Graphics, 0, 0, 0, 0, LockMousePage, argb(255, 255, 0, 0), near, 18, Bold
    
    GameDraw.Present TargetDC, 0, 0, 255
    
    If MouseState = 2 Then ResetClick: LockX = -1
    If MouseState = 0 And LockX <> -1 Then LockX = -1
    
    If DebugMode = True Then Exit Sub
    
CatchError:
    If Err.Number <> 0 Then
        Dim ErrNum As Long
        ErrNum = Err.Number
    '
    '┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
    '┃ * An error blocks the way !                                                      ┃
    '┃                                                                                              ┃
    '┃                                                                                              ┃
    '┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
    ' ERROR  LV 1                     20/20  █████████
    '┌────────┐  ┌────────┐   ┌────────┐  ┌────────┐
    '│     FIGHT    │  │       ACT     │   │      ITEM     │  │    MERCY   │
    '└────────┘  └────────┘   └────────┘  └────────┘
    '
    
        Dim temp As New Images
        temp.Create TargetDC, GWW, GWH
        BlurTo temp.CompatibleDC, 40
        BackImg.ImageByIndex(3).Present temp.CompatibleDC, 0, 0, 120
        GameFont.DrawText temp.Graphics, GWW / 2, GWH / 2 - 30, 0, 0, "不好意思 一个错误突然就从模块" & NowPage & "里跳出来了", argb(255, 255, 108, 105), center, 24, Bold
        GameFont.DrawText temp.Graphics, GWW / 2, GWH / 2, 0, 0, ErrNum & " " & FriendError(ErrNum), argb(180, 255, 108, 105), center, 16, Regular
        temp.Present TargetDC, 0, 0
        StopWork = True
    End If
End Sub
Private Sub Class_Initialize()
    ReDim Screens(0)
    Set GameCore = Me
    FadeDraw.Create TargetDC, GWW, GWH
End Sub
