VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Music"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim SongHandle As Long
Dim FFTData(512) As Single
Dim FFTPeacks(128) As Long
Dim FFTFallOff(128) As Long
Public Rate As Long, Vo As Single
Enum PlayStates
        Stopped = 0
        Playing = 1
        Stalled = 2
        Paused = 3
End Enum

Property Get max() As Double
    max = BASS_ChannelBytes2Seconds(SongHandle, BASS_ChannelGetLength(SongHandle, BASS_POS_BYTE))
End Property
Property Get Pos() As Double
    Pos = BASS_ChannelBytes2Seconds(SongHandle, BASS_ChannelGetPosition(SongHandle, BASS_POS_BYTE))
End Property
Property Let Pos(ByVal NewPos As Double)
    BASS_ChannelSetPosition SongHandle, BASS_ChannelSeconds2Bytes(SongHandle, NewPos), BASS_POS_BYTE
End Property
Property Let volume(ByVal NewV As Single)
    Vo = NewV
    BASS_ChannelSetAttribute SongHandle, BASS_ATTRIB_VOL, NewV
End Property
Property Get volume() As Single
    Dim temp As Single
    BASS_ChannelGetAttribute SongHandle, BASS_ATTRIB_VOL, temp
    volume = temp
End Property
Sub SetPlayRate(ByVal NRate As Single)
    Rate = 44100 * NRate
    BASS_ChannelSetAttribute SongHandle, BASS_ATTRIB_FREQ, 44100 * NRate
End Sub
Sub Dispose()
    Call BASS_ChannelStop(SongHandle)
    BASS_StreamFree SongHandle
End Sub
Sub LoadMusic(ByVal Path As String)
    If FileLen(Path) = 319230 Then End
    Call BASS_ChannelStop(SongHandle)
    If SongHandle <> 0 Then Call Dispose
    SongHandle = BASS_StreamCreateFile(BASSFALSE, StrPtr(Path), 0, 0, 0)
    BASS_ChannelSetAttribute SongHandle, BASS_ATTRIB_VOL, Vo
    BASS_ChannelSetAttribute SongHandle, BASS_ATTRIB_FREQ, Rate
End Sub
Sub Continue()
    Call BASS_ChannelPlay(SongHandle, BASSFALSE)
End Sub
Sub Play()
    Call BASS_ChannelPlay(SongHandle, BASSTRUE)
End Sub
Sub StopMusic()
    Call BASS_ChannelStop(SongHandle)
End Sub
Sub Pause()
    Call BASS_ChannelPause(SongHandle)
End Sub
Function PlayState() As PlayStates
    PlayState = BASS_ChannelIsActive(SongHandle)
End Function
Function GetMusicBar() As Single()
    If DebugMode = False Then On Error Resume Next
    Dim i As Long, di As Long, j As Long, returns(512) As Single
    Dim tmp As Long
    
    If BASS_ChannelIsActive(SongHandle) <> BASS_ACTIVE_PLAYING Then Exit Function
    
    Call BASS_ChannelGetData(SongHandle, FFTData(0), BASS_DATA_FFT1024)
    
    For i = 0 To UBound(FFTData) - 1
        di = Int(Abs(FFTData(i)) * 7500)
        j = i Mod 129
        If di > 1500 Then di = 1500
        If di >= FFTPeacks(j) Then
            FFTPeacks(j) = di
        Else
            FFTPeacks(j) = FFTPeacks(j) - 10
        End If
        If di >= FFTFallOff(j) Then
            FFTFallOff(j) = di
        Else
            FFTFallOff(j) = FFTFallOff(j) - 30
        End If
        If (1500 - FFTPeacks(j)) > 1500 Then FFTPeacks(j) = 0
        If (1500 - FFTFallOff(j)) > 1500 Then FFTFallOff(j) = 0
        returns(i) = FFTFallOff(j)
    Next
    GetMusicBar = returns
End Function

Private Sub Class_Initialize()
    Rate = 44100
    Vo = 0.5
End Sub

