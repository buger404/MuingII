VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images, Back As New Images
Dim LastPage As String, TitleIco As Integer, Text As String, OKButtons() As String
Dim OKResult As Integer
Public TipText As String, TipID As Integer
Public Property Get Visible() As Boolean
    Visible = True
End Property

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "Notify"
    LockMousePage = ""
    
    If mNowShow <> "Notify" Then GoTo DrawTip
    
    Back.Present Draw.CompatibleDC, 0, 0
    BackImg.ImageByIndex(3).Present Draw.CompatibleDC, 0, 0, 120
    
    MainImg.Image("dialog.png").Present Draw.CompatibleDC, GWW / 2 - 487 / 2, 180
    NotiImg.ImageByIndex(TitleIco).Present Draw.CompatibleDC, GWW / 2 - 487 / 2 + 20, 180 - 30
    
    GameFont.DrawText Draw.Graphics, GWW / 2 - 487 / 2 + 30, 180 + 40, 487 - 30 * 2, 275 - 30 - 40 - 50, Text, argb(255, 44, 41, 30), near, 16
    
    For i = 0 To UBound(OKButtons)
        CtrlImg.ImageByIndex(BMCtrl.BMNormalButton).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, GWW / 2 + 487 / 2 - 100 - i * 150, 420, OKButtons(i), argb(255, 255, 255, 255), 16, Regular, Button
        If IsClick = True Then
            OKResult = i: MouseState = 0
            GameCore.FadePage LastPage
        End If
    Next
    
    Draw.Present GameDraw.CompatibleDC, 0, 0
    
DrawTip:
    If TipText = "" Then
        If Ani.GetTarget(TipID) <> 0 Then Ani.ChangeTaskFast TipID, 0
    Else
        If Ani.GetTarget(TipID) <> 255 Then Ani.ChangeTaskFast TipID, 255
    End If
    If Ani.GetValue(TipID) = 0 Then Exit Sub
    
    Dim Y As Long
    
    With MainImg.Image("notifytip.png")
        Y = GWH - .Height
        .Present GameDraw.CompatibleDC, 0, Y, Ani.GetValue(TipID)
    End With
    
    GameFont.DrawText GameDraw.Graphics, 63, Y + 21, GWW - 40, 30, TipText, argb(Ani.GetValue(TipID), 255, 255, 255), near, 16, Regular
    
    TipText = ""

End Sub
Sub Message(ByVal nIcon As MuingII_Icons, nText As String, ParamArray Buttons())
    If mNowShow = "Notify" Then
        Do While mNowShow = "Notify"
            KeepDrawing
            DoEvents
        Loop
    End If
    OKResult = -1
    GameCore.FadePage "Notify"
    BlurTo Back.CompatibleDC, 20
    LastPage = mNowShow
    TitleIco = nIcon: Text = nText
    ReDim OKButtons(UBound(Buttons))
    For i = 0 To UBound(Buttons)
        OKButtons(i) = Buttons(i)
    Next
End Sub
Function Asks(ByVal nIcon As MuingII_Icons, nText As String, ParamArray Buttons()) As Integer
    If mNowShow = "Notify" Then
        Do While mNowShow = "Notify"
            KeepDrawing
            DoEvents
        Loop
    End If
    OKResult = -1
    GameCore.FadePage "Notify"
    BlurTo Back.CompatibleDC, 20
    LastPage = mNowShow
    TitleIco = nIcon: Text = nText
    ReDim OKButtons(UBound(Buttons))
    For i = 0 To UBound(Buttons)
        OKButtons(i) = Buttons(i)
    Next
    Do While OKResult = -1
        KeepDrawing
        DoEvents
    Loop
    Asks = OKResult
End Function

Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
    Back.Create TargetDC, GWW, GWH
    Set GameNotify = Me
    TipID = Ani.NewTask(500, 0, 0, 1)
End Sub



