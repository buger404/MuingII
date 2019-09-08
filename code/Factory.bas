Attribute VB_Name = "Factory"
Public Enum FontModes
    Regular = 0
    Bold = 1
    Italic = 2
End Enum
Public Enum PresentMode
    Normal = 0
    Central = 1
End Enum
Public Enum CtrlClass
    Button = 0
    ProgressBar = 1
    SliderBar = 2
    CheckBox = 3
    HScrollBar = 4
    VScrollBar = 5
    Button2 = 6
    EditBox = 7
    VScrollBar2 = 8
End Enum
Public Enum StrAlignment
    near = 0
    center = 1
    far = 2
End Enum
Dim mWW As Long, mWH As Long, mGameFrm As Form
Sub SetClickArea2(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    CtrlX = X: CtrlY = Y
    CtrlW = Width: CtrlH = Height
End Sub
Public Property Get WW() As Long
    WW = mWW
End Property
Public Property Get WH() As Long
    WH = mWH
End Property
Public Property Get GS() As Object
    Set GS = GameSave
End Property
Public Property Get NowShow() As String
    NowShow = mNowShow
End Property
Public Property Let NowShow(ByVal NewShow As String)
    mNowShow = NewShow
End Property
Public Sub PlayIcelolly(GameFrm As Object, GamePath As String, Power As Boolean)
    If Power = True Then
        InitGDIPlus
        Set mGameFrm = GameFrm
        BASS_Init -1, 44100, BASS_DEVICE_3D, GameFrm.hWnd, 0
        mWW = GameFrm.ScaleWidth: mWH = GameFrm.ScaleHeight
        GWW = mWW: GWH = mWH
        InitDustbin
        GameDraw.Create GameFrm.HDC, mWW, mWH
        TargetDC = GameFrm.HDC
        Dim GameScale As Single
        GameScale = 0.82
        BackImg.LoadDir GamePath & "\assets\background", GameScale
        MainImg.LoadDir GamePath & "\assets\main", GameScale
        PlayImg.LoadDir GamePath & "\assets\play", GameScale
        OnlineImg.LoadDir GamePath & "\assets\online", GameScale
        ModImg.LoadDir GamePath & "\assets\mod", GameScale
        CtrlImg.LoadDir GamePath & "\assets\controls", GameScale
        EditImg.LoadDir GamePath & "\assets\editor", GameScale
        PageImg.LoadDir GamePath & "\assets\page", GameScale
        NotiImg.LoadDir GamePath & "\assets\notify", GameScale
        BMImg.LoadDir GamePath & "\assets\bm", 0.3
        Set GameWin = GameFrm
        GameSave.Init
        Set GameBGM = New Music
        
        MainBack.Create TargetDC, GWW + 20, GWH + 20
        MainBack2.Create TargetDC, GWW + 20, GWH + 20
        MainBack3.Create TargetDC, GWW + 20, GWH + 20
        AutoSnd.LoadMusic App.Path & "\music\auto.wav"
        AutoSnd.volume = 0.7
        HitSnd.LoadMusic App.Path & "\music\hit.wav"
        HitSnd.volume = 0.3
        HitSnd2.LoadMusic App.Path & "\music\hit2.wav"
        HitSnd2.volume = 0.3
        MenuSnd.LoadMusic App.Path & "\music\menuhit.wav"
        MenuSnd.volume = 0.3
        MenuSnd2.LoadMusic App.Path & "\music\menuback.wav"
        MenuSnd2.volume = 0.5
        MissSnd.LoadMusic App.Path & "\music\miss.mp3"
        MissSnd.volume = 0.3
        WinSnd.LoadMusic App.Path & "\music\win.wav"
        WinSnd.volume = 0.5
        GdipGraphicsClear MainBack.Graphics, argb(255, 249, 249, 249)
    Else
        DoClearing
        TerminateGDIPlus
        BASS_Free
    End If
End Sub
Public Sub UpdateClickTest(ByVal X As Long, ByVal Y As Long, ByVal State As Long)
    MouseX = X: MouseY = Y
    MouseState = State
End Sub
Sub ResetClick()
    MouseState = 0
End Sub

