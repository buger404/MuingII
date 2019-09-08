Attribute VB_Name = "Core"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
           (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpDefault As Long, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function ChangeWindowMessageFilterEx Lib "user32" (ByVal hwnd As Long, ByVal Message As Long, ByVal action As Long, PCHANGEFILTERSTRUCT As Any) As Long
Public Declare Function ChangeWindowMessageFilter Lib "user32" (ByVal Message As Long, ByVal dwFlag As Long) As Long
Public Const MSGFLT_ALLOW = 1
Public Const MSGFLT_DISALLOW = 2
Public Const MSGFLT_RESET = 0
Public Const MSGFLT_ADD = 1
Public Const MSGFLT_REMOVE = 2
Public Const WM_COPYGLOBALDATA = 73
Public Const WM_DROPFILES = &H233

Dim KeepCTime As Long

Public HotAlpha As Long

Public Exp As Long

Public KeepSong As Boolean

Public UpdateUrl As String, UpdateVersion As String

Public LoginTime As String, GameOpenTime As Long, BaseTime As Single
Public ShowTask As Boolean, CLate As Long
Public Tasks As Tasks
Public MainPage As MainPage, EditPage As EditPage, Notify As Notify, ByePage As ByePage, SoloPage As SoloPage, PlayPage As PlayPage, DownPage As DownPage, SetPage As SetPage

Public MapCheck As String, MapList As String, MsgCheck As String
Public EditBoxs(999) As Boolean, EditBoxText(999) As String, LastBox As Long, NowEdit As Long
Public UserID As Long, UserName As String
Public PageImg As New ImageCollection
Public BackImg As New ImageCollection, MainImg As New ImageCollection, CtrlImg As New ImageCollection, PlayImg As New ImageCollection, ModImg As New ImageCollection, OnlineImg As New ImageCollection, EditImg As New ImageCollection, BMImg As New ImageCollection, NotiImg As New ImageCollection
Public GameFont As Fonts, GameDraw As New Images, TargetDC As Long, GameWin As Form, GameSave As New Saving, GameBGM As Music
Public GameCore As GameManager, GameNotify As Notify
Public MouseX As Long, MouseY As Long, MouseState As Long
Public CtrlX As Long, CtrlY As Long, CtrlW As Long, CtrlH As Long
Public LockX As Long, LockY As Long, LockW As Long, LockH As Long
Public GWW As Long, GWH As Long
Public mNowShow As String
Public LeaveCount As Single
Public HighMode As Boolean
Public UnClicked As Boolean
Public MainBack As New Images, MainBack2 As New Images, MainBack3 As New Images, BackChangeTime As Currency
Public HitSnd As New Music, MissSnd As New Music, MenuSnd As New Music, MenuSnd2 As New Music, WinSnd As New Music, HitSnd2 As New Music, AutoSnd As New Music
Public SongPath As String, NowSong As Integer
Public SongList() As SongListT
Public MouseImage As Integer
Public LastDrawTime As Currency
Public CloseTime As Boolean
Public NowSongD As Integer

Public LockMousePage As String, NowPage As String

Public DrawUse As Currency, DrawUse2 As Currency
Public DrawCount As Long

Public Sock As Winsock, Connected As Boolean, Newest As Boolean, GiveupConnect As Boolean

Public ModPower(3) As Boolean

Public Ani As New Animation

Public VC1 As Long, VC2 As Long, VC3 As Long
Public TC1 As Long, TC2 As Long, TC3 As Long
Public AC1 As Long, AC2 As Long, AC3 As Long

Public Settings(7) As Integer
Public HistoryFPS(20) As Long, HistoryDraw(20) As Long, LastDrawC As Long

Public Type AniTask
    StartTime As Long
    During As Long
    Target As Single
    Oran As Single
    Types As Integer
End Type
Public Enum MuingII_Settings
    MuingShowHot = 1
    MuingShowHot2 = 2
    MuingBlurBack = 3
    MuingUseDog = 4
    MuingNoServer = 5
    MuingDebug = 6
    MuingShowAllLines = 7
End Enum
Public Enum MuingII_Icons
    MuingNoConnected = 1
    MuingServer = 2
    MuingUpdate = 3
    MuingError = 4
    MuingAsk = 5
    MuingInformation = 6
    MuingSuccess = 7
End Enum
Public Enum BMCtrl
    BMNormalButton = 1
    BMVScroll = 2
End Enum
Public Enum MuingII_Object
    ObjNormal = 0
    ObjLongHit = 1
End Enum
Public Type MuingII_MapData
    PressCheck(3) As Boolean
    PressType(3) As Long
    PressTime(3) As Long
    StartTime As Long
End Type
Public Type MuingII_Map
    MapData() As MuingII_MapData
    ObjSpeed As Single
End Type
Public Type MuingII_MapFile
    Levels(2) As MuingII_Map
    Title As String
    Maker As String
    Artist As String
    MapID As Long
    MapVersion As Long
    DownloadSource As String
    MuingVersion As Long
End Type
Public Type SongListT
    MODs(2) As String
    Accuracy(2) As Single
    SongPic As Images
    SongPicCircle As Images
    Path As String
    Info As MuingII_MapFile
    Difficulty(2) As Single
    Grade(2) As String
    Score(2) As Long
    MaxCombo(2) As Long
    Rank(2) As String
End Type
Public EditMap As MuingII_MapFile

Public MapCheckOK As Boolean, MapCheckID As Integer
Sub ChangeSettings(ByVal Index As Integer, ByVal Val As Integer)
    Settings(Index) = Val
    GameSave.WSave "Settings" & Index, Val
End Sub
Sub Send(ByVal Text As String)
    Sock.SendData Text & Chr(404)
End Sub
Sub KeepDrawing()
    Call MainWindow.DrawGame
    If Connected Then
        If GetTickCount - KeepCTime > 2000 Then
            KeepCTime = GetTickCount
            Send "t*" & GetTickCount
        End If
    End If
End Sub
Function CheckText(Text As String) As Boolean
    CheckText = (InStr(Text, "*") > 0 Or InStr(Text, "/") > 0 Or InStr(Text, "\") > 0 Or InStr(Text, "|") > 0 Or InStr(Text, "?") > 0 Or InStr(Text, """") > 0 Or InStr(Text, "<") > 0 Or InStr(Text, ">") > 0 Or InStr(Text, ":") > 0)
End Function
Function LoadMapFile(ByVal Path As String) As Boolean
    On Error GoTo sth
    
    Open App.Path & "\LoadLog.txt" For Append As #2
    Print #2, Now & "    命令行：" & Command
    Close #2
    
    Dim FSO As Object, tempMap As MuingII_MapFile
    Dim ESongList As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    For i = 1 To 1
    
        If Dir(App.Path & "\temp\", vbDirectory) <> "" Then FSO.DeleteFolder App.Path & "\temp"
        
        Export Path, App.Path & "\temp\"
        
        Open App.Path & "\temp\map.mu" For Binary As #1
        Get #1, , tempMap
        Close #1
        
        Export Path, SongPath & "\" & tempMap.MapID & " - " & tempMap.Title & " - " & tempMap.Maker
        ESongList = ESongList & tempMap.MapID & " - " & tempMap.Title & " - " & tempMap.Maker & " - " & tempMap.Artist & vbCrLf
        
        If Dir(App.Path & "\temp\", vbDirectory) <> "" Then FSO.DeleteFolder App.Path & "\temp"
    Next
        
    LoadMapFile = True
    
sth:
    If Err.Number <> 0 Then
        Open App.Path & "\LoadLog.txt" For Append As #2
        Print #2, Now & "   导入谱子失败。 " & Err.Number & "：" & Err.Description
        Close #2
    End If
End Function
Sub Main()
    Dim sTmp As String * 255, nLength As Long, pidl As Long
    SHGetSpecialFolderLocation 0, &H5, pidl
    SHGetPathFromIDList pidl, sTmp
    SongPath = Left(sTmp, InStr(sTmp, Chr(0)) - 1) & "\Muing2"
    If Dir(SongPath, vbDirectory) = "" Then MkDir SongPath

    If Command <> "" Then
        On Error GoTo sth
        
        Open App.Path & "\LoadLog.txt" For Append As #2
        Print #2, Now & "    命令行：" & Command
        Close #2
        
        Dim data(1) As String
        Dim FSO As Object, tempMap As MuingII_MapFile, Path As String
        Dim ESongList As String
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        data(1) = Replace(Command, """", "")
        
        For i = 1 To 1
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
        
        If Val(GetSetting("MuingII", "RunTime", "Window")) <> 0 Then
            SendMessageA Val(GetSetting("MuingII", "RunTime", "Window")), 404, 0, 0
        End If
        
        MsgBox "导入成功" & vbCrLf & ESongList, 64, "嘤悦2"
        
sth:
        If Err.Number <> 0 Then
            MsgBox "导入失败。", 16, "嘤悦2"
            Open App.Path & "\LoadLog.txt" For Append As #2
            Print #2, Now & "   导入谱子失败。 " & Err.Number & "：" & Err.Description
            Close #2
        End If
        End
    End If
    
    On Error Resume Next
    MainWindow.Show
    Sock.LocalPort = Int(Rnd * 1000 + 20) + 4046
    
    Do While CloseTime = False
        KeepDrawing
        DoEvents
    Loop
    
    If DebugMode = False Then SetWindowLongA MainWindow.hwnd, GWL_WNDPROC, MWndProc
    
    On Error Resume Next
    Sock.close
    
    PlayIcelolly MainWindow, App.Path, False
    Unload MainWindow
End Sub
Function CheckServer(Optional CheckLogin As Boolean = False, Optional NoNotify As Boolean = False) As Boolean
    If Connected = True Then
        If Newest = True Then
            If CheckLogin = True Then
                If UserID = 0 Then
                    CheckServer = False
                    GameNotify.Message MuingServer, "请先登陆您的账号", "噢"
                Else
                End If
            End If
        Else
        End If
    Else
        If GiveupConnect Then
        Else
        End If
    End If
End Function
Sub PlayRandomSong()
    Randomize
    Dim id As Long
    id = Int(Rnd * UBound(SongList))
    If id = 0 Then id = 1
    GameBGM.LoadMusic SongPath & "\" & SongList(id).Path & "\music.mp3"
    GameBGM.Play
    SetMainBack SongPath & "\" & SongList(id).Path & "\background.png"
    SoloPage.ChangeIndex = NowSong: SoloPage.ChangeTime = GetTick
    SoloPage.ScrollTo id
    NowSong = id
End Sub
Function GetDifficulty(map As MuingII_Map) As Single
    Dim ret As Single
    Dim NoteLine As Single, NoteTime As Single, Pos As Integer, LastPos As Integer
    
    For i = 1 To UBound(map.MapData)
        Pos = 0
        
        If map.MapData(i).PressCheck(0) Then Pos = 0
        If map.MapData(i).PressCheck(1) Then Pos = 1
        If map.MapData(i).PressCheck(2) Then Pos = 2
        If map.MapData(i).PressCheck(3) Then Pos = 3
        
        NoteLine = Abs(Pos - LastPos) / 2
        NoteTime = Abs((map.MapData(i).StartTime - map.MapData(i - 1).StartTime) / 1000) / 4400
        If (map.MapData(i).StartTime - map.MapData(i - 1).StartTime) < 40 Then NoteTime = 1
        ret = ret + NoteLine / NoteTime + 5 / NoteTime + NoteLine * 2
        
        LastPos = Pos
    Next
    
    ret = ret * (1 + (map.ObjSpeed / 10))
    
    ret = ret / (UBound(map.MapData) + 1)
    ret = Int(ret / 100) / 130

    GetDifficulty = ret
End Function
Function ColorLight(ByVal r As Long, ByVal G As Long, ByVal b As Long) As Long
Dim result As Long
If r > result Then result = r
If G > result Then result = G
If b > result Then result = b
ColorLight = (result - r) ^ 2 + (result - G) ^ 2 + (result - b) ^ 2
End Function
Sub SetMainBack(Path As String)
    MainBack.Present MainBack2.CompatibleDC, 0, 0

    Dim Image As Long, Width As Long, Height As Long, Color As Long, ColorA(3) As Byte, Light As Single
    CreateImage StrPtr(Path), Image
    
    GdipGetImageWidth Image, Width: GdipGetImageHeight Image, Height
    If (Width <> 1088 Or Height <> 723) And mNowShow = "EditPage" Then GameNotify.Message MuingInformation, "背景图片分辨率为1088x723最佳哟。", "继续"
    If Settings(MuingII_Settings.MuingBlurBack) = 0 Then BlurImage Image, Width, Height, 30
    GdipDrawImageRect MainBack.Graphics, Image, 0, 0, 3, 3
    Color = GetPixel(MainBack.CompatibleDC, 2, 2)
    
    CopyMemory ColorA(0), Color, 4
    
    Light = (ColorA(0) / 255 + ColorA(1) / 255 + ColorA(2) / 255) / 3

    VC1 = 127 + 127 * IIf(Light < 0.5, Light, -(Light - 0.5)): ControlLong VC1, 0, 255
    VC2 = 127 + 127 * IIf(Light < 0.5, Light, -(Light - 0.5)): ControlLong VC2, 0, 255
    VC3 = 127 + 127 * IIf(Light < 0.5, Light, -(Light - 0.5)): ControlLong VC3, 0, 255
    
    TC1 = VC1 + 50 * IIf(Light < 0.5, 1, -1): ControlLong TC1, 0, 255
    TC2 = VC2 + 50 * IIf(Light < 0.5, 1, -1): ControlLong TC2, 0, 255
    TC3 = VC3 + 50 * IIf(Light < 0.5, 1, -1): ControlLong TC3, 0, 255
    
    'AC1 = ColorA(0) + 100 * IIf(Light < 0.5, 1, -1): ControlLong AC1, 0, 255
    'AC2 = ColorA(1) + 100 * IIf(Light < 0.5, 1, -1): ControlLong AC2, 0, 255
    'AC3 = ColorA(2) + 100 * IIf(Light < 0.5, 1, -1): ControlLong AC3, 0, 255
    AC1 = 0: AC2 = 234: AC3 = 118
    
    'If Light < 0.5 Then
        'VC1 = 242: VC2 = 242: VC3 = 242
    'Else
        'VC1 = 27: VC2 = 27: VC3 = 27
    'End If
    
    GdipGraphicsClear MainBack.Graphics, argb(255, 249, 249, 249)
    GdipDrawImageRect MainBack.Graphics, Image, 0, 0, GWW + 20, GWH + 20
    DelImage Image
    
    MainBack.Present MainBack3.CompatibleDC, 0, 0
    BackChangeTime = GetTick
    
End Sub
Sub ControlLong(value As Long, min As Long, max As Long)
    If value < min Then value = min
    If value > max Then value = max
End Sub
Public Sub CircleImage(Image As Long, r As Long)
    Dim mClipPath As Long, bmpGraph As Long, BMP As Long, mTexture As Long, bmpGraph2 As Long, BMP2 As Long
    Dim Width As Long, Height As Long, r2 As Long
    Dim BMP3 As Long, bmpGraph3 As Long
    
    GdipGetImageWidth Image, Width: GdipGetImageHeight Image, Height
    r2 = IIf(Width > Height, Height, Width)
    GdipCreateBitmapFromScan0 r, r, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP
    GdipCreateBitmapFromScan0 Width, Height, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP2
    GdipCreateBitmapFromScan0 r2, r2, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP3
    GdipGetImageGraphicsContext BMP3, bmpGraph3
    GdipGetImageGraphicsContext BMP, bmpGraph
    GdipGetImageGraphicsContext BMP2, bmpGraph2
    
    GdipCreatePath FillModeWinding, mClipPath
    GdipAddPathEllipse mClipPath, Width / 2 - r2 / 2, Height / 2 - r2 / 2, r2 - 1, r2 - 1
    
    GdipCreateTexture Image, WrapModeClamp, mTexture
    GdipSetSmoothingMode bmpGraph, SmoothingModeAntiAlias: GdipSetSmoothingMode bmpGraph2, SmoothingModeAntiAlias: GdipSetSmoothingMode bmpGraph3, SmoothingModeAntiAlias
    GdipFillPath bmpGraph2, mTexture, mClipPath
    GdipDrawImage bmpGraph3, BMP2, -(Width / 2 - r2 / 2), -(Height / 2 - r2 / 2)
    GdipDrawImageRect bmpGraph, BMP3, 0, 0, r, r

    DelImage Image: Image = BMP
    GdipDeleteGraphics bmgraph: GdipDeleteGraphics bmgraph2: DelImage BMP2
    GdipDeleteGraphics bmgraph3: DelImage BMP3
End Sub
Sub GetSongList()
    For i = 1 To UBound(SongList)
        SongList(i).SongPic.Dispose
        SongList(i).SongPicCircle.Dispose
        Set SongList(i).SongPic = Nothing
        Set SongList(i).SongPicCircle = Nothing
    Next
    ReDim SongList(0)
    Dim Folder As String, Image As Long, Width As Long, Height As Long, temp As MuingII_MapFile
    Dim temp2() As String
    Dim ErrFiles As String
    Folder = Dir(SongPath & "\", vbDirectory)
    Do While Folder <> ""
        DoEvents
        If Folder <> "." And Folder <> ".." Then
            Open SongPath & "\" & Folder & "\map.mu" For Binary As #1
            Get #1, , temp
            Close #1
            
            If temp.MapID = 0 Then ErrFiles = ErrFiles & Folder & vbCrLf: GoTo last
            
            temp2 = Split(Folder, " - ")
            If UBound(temp2) <> 2 Then ErrFiles = ErrFiles & Folder & vbCrLf: GoTo last
            
            If Val(temp.MuingVersion) < 1 Then
                GameNotify.Message MuingError, "无法载入该谱子，因为它太旧了。（版本号：" & IIf(Val(temp.MuingVersion) = 0, "未知", temp.MuingVersion) & "）" & vbCrLf & Folder, "继续"
            Else
                ReDim Preserve SongList(UBound(SongList) + 1)
                SongList(UBound(SongList)).Path = Folder
                Set SongList(UBound(SongList)).SongPic = New Images
                Set SongList(UBound(SongList)).SongPicCircle = New Images
                SongList(UBound(SongList)).SongPic.Create TargetDC, 575, 104
                SongList(UBound(SongList)).SongPicCircle.Create TargetDC, 128, 128
                
                CreateImage StrPtr(SongPath & "\" & Folder & "\background.png"), Image
                If Image <> 0 Then
                    GdipGetImageWidth Image, Width: GdipGetImageHeight Image, Height
                    GdipDrawImageRect SongList(UBound(SongList)).SongPic.Graphics, Image, 0, 104 / 2 - Height / (Width / 575) / 2, 575, Height / (Width / 575)
                    
                    CircleImage Image, 128
                    GdipDrawImage SongList(UBound(SongList)).SongPicCircle.Graphics, Image, 0, 0
        
                    DelImage Image
                Else
                    GdipGraphicsClear SongList(UBound(SongList)).SongPicCircle.Graphics, argb(255, 249, 249, 249)
                End If
                
                SongList(UBound(SongList)).Info = temp
                
                For i = 0 To 2
                    SongList(UBound(SongList)).Difficulty(i) = GetDifficulty(SongList(UBound(SongList)).Info.Levels(i))
                    SongList(UBound(SongList)).Grade(i) = GameSave.RSave(temp.Title & "." & temp.Maker & "." & temp.MapID & ".Grade" & i)
                    If SongList(UBound(SongList)).Grade(i) = "" Then SongList(UBound(SongList)).Grade(i) = "-"
                    SongList(UBound(SongList)).MaxCombo(i) = Val(GameSave.RSave(temp.Title & "." & temp.Maker & "." & temp.MapID & ".Combo" & i))
                    SongList(UBound(SongList)).Score(i) = Val(GameSave.RSave(temp.Title & "." & temp.Maker & "." & temp.MapID & ".Score" & i))
                    SongList(UBound(SongList)).Accuracy(i) = Val(GameSave.RSave(temp.Title & "." & temp.Maker & "." & temp.MapID & ".Accuracy" & i))
                    SongList(UBound(SongList)).MODs(i) = GameSave.RSave(temp.Title & "." & temp.Maker & "." & temp.MapID & ".MOD" & i)
                    If SongList(UBound(SongList)).MODs(i) = "" Then SongList(UBound(SongList)).MODs(i) = "- "
                    SongList(UBound(SongList)).Rank(i) = GameSave.RSave(temp.Title & "." & temp.Maker & "." & temp.MapID & ".Rank" & i)
                    If SongList(UBound(SongList)).Rank(i) = "" Then SongList(UBound(SongList)).Rank(i) = " - "
                Next
            End If
last:
        End If
        Folder = Dir(, vbDirectory)
    Loop
    
    If ErrFiles <> "" Then GameNotify.Message MuingError, "为配合多人游戏，由于以下旧版谱子没有正确的谱子ID，未能载入。" & vbCrLf & ErrFiles, "fuck"
End Sub
Function CheckFileHeader(Path As String, header As String) As Boolean
    CheckFileHeader = True
    Dim data As Byte
    Open Path For Binary As #1
    For i = 1 To Len(header)
        Get #1, i, data
        If Chr(data) <> Mid(header, i, 1) Then CheckFileHeader = False
    Next
    Close #1
End Function
Function cubicCurves(t As Single, value0 As Single, value1 As Single, Value2 As Single, value3 As Single) As Single
    cubicCurves = (value0 * ((1 - t) ^ 3)) + (3 * value1 * t * ((1 - t) ^ 2)) + (3 * Value2 * (t ^ 2) * (1 - t)) + (value3 * (t ^ 3))
    '贝塞尔曲线公式： B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
End Function
Function GetLongTime() As String
    GetLongTime = year(Now) & format(Month(Now), "00") & format(Day(Now), "00") & format(Hour(Now), "00") & format(Minute(Now), "00") & format(Second(Now), "00")
End Function
'MouseState: 0-None 1-Down 2-Up
Function IsMouseIn() As Boolean
    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH Then IsMouseIn = True
        Exit Function
    End If
    
    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        IsMouseIn = True
        MouseImage = 1
    End If
End Function
Function IsClick() As Boolean
    If LockMousePage <> "" Then
        If NowPage <> LockMousePage Then Exit Function
    End If
    
    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH And MouseState = 2 Then
            LockX = -1
            IsClick = True
        End If
        Exit Function
    End If

    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        If MouseState = 2 Then
            LockX = -1
            IsClick = True
        End If
    End If
    If UnClicked Then IsClick = False
End Function
Function IsMouseDown() As Boolean
    If LockMousePage <> "" Then
        If NowPage <> LockMousePage Then Exit Function
    End If

    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH And MouseState = 1 Then IsMouseDown = True
        Exit Function
    End If
    
    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        If MouseState = 1 Then
            IsMouseDown = True
            LockX = CtrlX: LockY = CtrlY: LockW = CtrlW: LockH = CtrlH
        End If
    End If
    If UnClicked Then IsMouseDown = False
End Function
Function IsMouseDownNoKeep() As Boolean
    If LockMousePage <> "" Then
        If NowPage <> LockMousePage Then Exit Function
    End If

    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH And MouseState = 1 Then IsMouseDownNoKeep = True
        Exit Function
    End If
    
    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        If MouseState = 1 Then
            IsMouseDownNoKeep = True
        End If
    End If
    If UnClicked Then IsMouseDownNoKeep = False
End Function
Function IsMouseUp() As Boolean
    
    If LockMousePage <> "" Then
        If NowPage <> LockMousePage Then Exit Function
    End If

    If LockX <> -1 Then
        If LockX = CtrlX And LockY = CtrlY And LockW = CtrlW And LockH = CtrlH And MouseState = 2 Then
            LockX = -1
            IsMouseUp = True
        End If
        Exit Function
    End If
    
    If MouseX >= CtrlX And MouseY >= CtrlY And MouseX <= CtrlX + CtrlW And MouseY <= CtrlY + CtrlH Then
        If MouseState = 2 Then
            LockX = -1
            IsMouseUp = True
        End If
    End If
    If UnClicked Then IsMouseUp = False
End Function
Sub CreateFolder(ByVal Path As String)
    Dim temp() As String, NowPath As String
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    temp = Split(Path, "\")
    For i = 0 To UBound(temp) - 1
        If temp(i) Like "*.*" Then Exit Sub
        NowPath = NowPath & temp(i) & "\"
        If Dir(NowPath, vbDirectory) = "" Then MkDir NowPath
    Next
End Sub
Public Sub BlurTo(dc As Long, Optional Radius As Long = 60)
    If GameWin Is Nothing Then Exit Sub
    Dim Image As Long, Graphics As Long
    GameWin.AutoRedraw = True
    GameDraw.Present GameWin.hdc, 0, 0
    GameWin.Refresh
    DoEvents
    GdipCreateBitmapFromHBITMAP GameWin.Image.Handle, GameWin.Image.hpal, Image
    BlurImage Image, GWW, GWH, Radius
    GdipCreateFromHDC dc, Graphics
    GdipDrawImage Graphics, Image, 0, 0
    DelImage Image
    GdipDeleteGraphics Graphics
    GameWin.AutoRedraw = False
End Sub
Sub BlurImage(Image As Long, Width As Long, Height As Long, Optional Radius As Long = 60)
    Dim Effect As Long
    Dim p As BlurParams
    GdipCreateEffect2 GdipEffectType.Blur, Effect
    p.Radius = Radius
    GdipSetEffectParameters Effect, p, LenB(p)
    GdipBitmapApplyEffect Image, Effect, NewRectL(0, 0, Width, Height), 0, 0, 0
    GdipDeleteEffect Effect
End Sub
Public Function ReadINI(ByVal SectionName As String, ByVal KeyName As String, ByVal IniFileName As String) As String
    Dim strBuf As String
    strBuf = String(128, 0)
    GetPrivateProfileString StrPtr(SectionName), StrPtr(KeyName), StrPtr(""), StrPtr(strBuf), 128, StrPtr(IniFileName)
    strBuf = Replace(strBuf, Chr(0), "")
    ReadINI = strBuf
End Function
Public Function FriendError(ByVal Num As Long) As String
    FriendError = Error(Num)
    Select Case Num
        Case 0
        FriendError = "等一下？哪来的错误？"
        Case 5
        FriendError = "多半是该死的404又忘记把废弃的代码删干净了。"
        Case 6
        FriendError = "shit 404为变量倒果汁的时候不小心溢出来了。"
        Case 7
        FriendError = "真的很抱歉，404没有做好清洁工的职责。"
        Case 9
        FriendError = "404让我在数组的边缘试探...哦豁，掉入悬崖了。"
        Case 11
        FriendError = "对不起，脑残404数学不过关，除以0什么的..."
        Case 13
        FriendError = "404刚才在给别人介绍对象的时候被双方左右各扇了一巴掌。"
        Case 28
        FriendError = "罢工！404一次性让我们做太多的工作了！"
        Case 35
        FriendError = "多半是该死的404又忘记把废弃的代码删干净了。"
        Case 52
        FriendError = "404给错了房间号码。"
        Case 53
        FriendError = "等等...404的GPS出了点问题..."
        Case 55
        FriendError = "丢三落四的404打开一个文件后忘了关上。"
        Case 58
        FriendError = "嗯。。这个文件已经存在了，用TNT炸掉么？"
        Case 70
        FriendError = "抱歉，身在底层的我实在没有权利完成这项工作。"
        Case 75
        FriendError = "404！这个地址是什么毛线啦！"
        Case 76
        FriendError = "404！这里已经拆迁了啦！"
    End Select
End Function
