Attribute VB_Name = "WndProc"
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As Long, ByVal fAccept As Long)
Public Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long

Public MWndProc As Long, SafeLock As Boolean

Public Function FunWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo sth
    If SafeLock Then GoTo Last
    
    If uMsg = WM_MOUSEWHEEL Then
        Dim Direction As Integer, Strong As Single
        Direction = IIf(wParam < 0, 1, 0): Strong = Abs(wParam / 7864320)
        If mNowShow = "SoloPage" Then SoloPage.MouseWheel Direction, Strong
    End If
    
    If uMsg = 404 Then
        GetSongList
        If UBound(SongList) > 0 Then
            PlayRandomSong
        End If
    End If
    
    If uMsg = WM_DROPFILES Then
        Dim hDrop As Long, nLoopCtr As Integer, IReturn As Long, sFileName As String
        Dim FileList As String
        hDrop = wParam: sFileName = Space$(255)
        nDropCount = DragQueryFile(hDrop, -1, sFileName, 254)
        For nLoopCtr = 0 To nDropCount - 1
            sFileName = Space$(255)
            IReturn = DragQueryFile(hDrop, nLoopCtr, sFileName, 254)
            FileList = FileList & Left$(sFileName, IReturn) & vbCrLf
        Next
        Call DragFinish(hDrop)
        MainWindow.FileDrop FileList
    End If
    
Last:
    FunWndProc = CallWindowProc(MWndProc, hWnd, uMsg, wParam, lParam)
    
    Exit Function

sth:
    SafeLock = True: GoTo Last
End Function
