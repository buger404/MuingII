Attribute VB_Name = "GetTime"
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public CPURTMS As Double, CPURTS As Currency
Sub GetCPURTS()
    If QueryPerformanceFrequency(CPURTS) Then
        CPURTMS = CPURTS / 1000
    Else
        'Debug.Print "Failed to init ."
    End If
End Sub
Function GetTick() As Currency
    If CPURTS = 0 Then GetCPURTS
    
    Dim temp As Currency
    QueryPerformanceCounter temp
    GetTick = temp
 End Function
 
Function GetPastTick(Tick As Currency) As Currency
    If CPURTS = 0 Then GetCPURTS

    Dim temp As Currency
    QueryPerformanceCounter temp
    GetPastTick = (temp - Tick) * 1000 / CPURTS
 End Function
