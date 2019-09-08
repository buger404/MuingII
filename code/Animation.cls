VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim AnimateTasks() As AniTask

Sub ChangeTaskFast(ByVal AniID As Integer, ByVal Target As Single)

    AnimateTasks(AniID).Oran = AnimateTasks(AniID).Target
    AnimateTasks(AniID).Target = Target
    AnimateTasks(AniID).StartTime = GetTickCount
    
End Sub

Sub ChangeTask(ByVal AniID As Integer, ByVal During As Long, ByVal Target As Single, ByVal Oran As Single, ByVal Types As Integer)

    AnimateTasks(AniID).During = During
    AnimateTasks(AniID).Oran = Oran
    AnimateTasks(AniID).Target = Target
    AnimateTasks(AniID).StartTime = GetTickCount
    AnimateTasks(AniID).Types = Types
    
End Sub

Function NewTask(ByVal During As Long, ByVal Target As Single, ByVal Oran As Single, ByVal Types As Integer) As Integer
    ReDim Preserve AnimateTasks(UBound(AnimateTasks) + 1)
    AnimateTasks(UBound(AnimateTasks)).During = During
    AnimateTasks(UBound(AnimateTasks)).Oran = Oran
    AnimateTasks(UBound(AnimateTasks)).Target = Target
    AnimateTasks(UBound(AnimateTasks)).StartTime = GetTickCount
    AnimateTasks(UBound(AnimateTasks)).Types = Types
    
    NewTask = UBound(AnimateTasks)
End Function

Function GetTarget(ByVal AniID As Integer) As Single
    GetTarget = AnimateTasks(AniID).Target
End Function

Function GetValue(ByVal AniID As Integer) As Single
    Dim Pro As Single
    Pro = (GetTickCount - AnimateTasks(AniID).StartTime) / AnimateTasks(AniID).During
    If Pro > 1 Then Pro = 1
    If AnimateTasks(AniID).Types = 1 Then Pro = cubicCurves(Pro, 0, 0, 1, 1) 'Ease
    If AnimateTasks(AniID).Types = 2 Then Pro = cubicCurves(Pro, 0, 1, 1, 1) 'FadeIn
    If AnimateTasks(AniID).Types = 3 Then Pro = cubicCurves(Pro, 0, 0, 0, 1) 'FadeOut
    
    GetValue = AnimateTasks(AniID).Oran + (AnimateTasks(AniID).Target - AnimateTasks(AniID).Oran) * Pro

End Function

Private Sub Class_Initialize()
    ReDim AnimateTasks(0)
End Sub
