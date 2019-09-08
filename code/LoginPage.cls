VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoginPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images, Mode As Long
Dim Remember As Boolean
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "LoginPage")
End Property

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    Call DrawPage0
    
    SetClickArea2 GWW - 210, -20, 210, 146
    MainImg.Image("backbutton.png").Present Draw.CompatibleDC, GWW - 210, -20, IIf(IsMouseIn, 255, 170)
    If IsClick = True Then
        GameCore.ChangePage "MainPage"
    End If
    
    Draw.Present GameDraw.CompatibleDC, 0, 0
End Sub
Public Sub DrawPage0()
    MainBack.Present Draw.CompatibleDC, MouseX / GWW * -20, MouseY / GWH * -20
    
    Dim Width2 As Long, Height2 As Long, Deep As Long, Path As Long, Brush As Long, Ox As Long
    Dim Y As Long, X As Long, TargetI As Integer
    Width2 = GWW: Height2 = GWH: Deep = 350: Ox = GetTickCount / 5
    TargetI = Int((220 + 64) / 40) * 40
    
    GdipCreateLineBrush NewPointF(GWW, 0), NewPointF(GWW, GWH), argb(255, 162, 89, 255), argb(40, TC1, TC2, TC3), WrapModeTile, Brush
    GdipCreatePath FillModeWinding, Path
    For i = 0 To GWW + 100 Step 40
        X = i + Ox
        GdipAddPathLine Path, X - Ox, Sin(X / 150) * 30 + Height2 - Deep - 50, X - Ox, Sin(X / 150) * 30 + Height2 - Deep - 50
        If i = TargetI Then Y = Sin(X / 150) * 30 + Height2 - Deep - 50
    Next
    GdipAddPathLine Path, Width2, Height2, 0, Height2

    GdipFillPath Draw.Graphics, Brush, Path

    GdipDeletePath Path
    GdipDeleteBrush Brush

    Y = Y - 130

    GameFont.DrawText Draw.Graphics, 195, Y + 205, 100, 30, IIf(Mode = 0, "Login", "Join us"), argb(255 - (Y + 205) / GWH * 255, 249, 249, 249), Near, 16, Bold
    If Mode = 0 Then CtrlImg.ImageByIndex(3).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 195, Y + 250, 1, argb(255, 249, 249, 249), 16, Regular, EditBox
    If EditBoxText(2) = "password" Then
        CtrlImg.ImageByIndex(3).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 195, IIf(Mode = 0, Y + 295, Y + 250), 2, argb(255, 249, 249, 249), 16, Regular, EditBox
    Else
        CtrlImg.ImageByIndex(3).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 195, IIf(Mode = 0, Y + 295, Y + 250), 2, argb(255, 249, 249, 249), 16, Regular, EditBox, , 1
    End If
    
    If EditBoxText(1) <> "userid" Then
        EditBoxText(1) = Int(Val(EditBoxText(1)))
        If Len(EditBoxText(1)) > 5 Then EditBoxText(1) = Left(EditBoxText(1), 5)
    End If
    
    If Mode = 0 Then
        CtrlImg.ImageByIndex(4).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 225, Y + 350, "记住密码并自动登陆", argb(255, 255, 255, 255), 16, Regular, CheckBox, , Remember
    End If
    
    GameFont.DrawText Draw.Graphics, 658, Y + 338, 100, 30, IIf(Mode = 0, "注册一个", "返回登陆"), argb(255, 249, 249, 249), Near, 16
    SetClickArea2 658, Y + 338, 100, 30
    If IsClick Then
        If Mode = 0 Then
            Mode = 1: EditBoxText(2) = "password"
        Else
            Mode = 0
        End If
    End If
    
    CtrlImg.ImageByIndex(BMCtrl.BMNormalButton).PresentWithCtrl Draw.CompatibleDC, Draw.Graphics, 605, Y + 420, IIf(Mode = 0, "登陆", "注册"), argb(255, 255, 255, 255), 14, Regular, Button
    If IsClick Then
        If EditBoxText(2) = "" Then GameNotify.Message "= =", "密码不能为空。", "噢": Exit Sub
        If Mode = 0 Then
            Send "log*" & Int(Val(EditBoxText(1))) & "*" & EditBoxText(2)
        Else
            Send "reg*" & EditBoxText(2)
        End If
        Sleep 1000
        MsgCheck = ""
        Do While MsgCheck = ""
            DoEvents
        Loop
        Dim temp() As String
        If Mode = 0 Then
            temp = Split(MsgCheck, "*")
            If temp(1) = "no" Then
                GameNotify.Message "(╯‵□′)╯︵┻━┻", "用户名或密码不正确！", "噫？"
            Else
                If Remember Then GameSave.WSave "UserID", Int(Val(EditBoxText(1))): GameSave.WSave "Password", EditBoxText(2)
                mNowShow = "MainPage"
                GameNotify.Message ">UO", "登陆成功", "好"
                UserName = temp(1): UserID = Int(Val(EditBoxText(1)))
            End If
        Else
            temp = Split(MsgCheck, "*")
            mNowShow = "MainPage"
            GameNotify.Message "OVO", "注册成功，请牢记您的账号ID：" & vbCrLf & temp(1), "OK"
            UserName = temp(2): UserID = Int(Val(temp(1)))
        End If
    End If
End Sub

Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
End Sub




