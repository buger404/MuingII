VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Saving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Muing2_Save
    Title As String
    value As String
End Type
Private Type Muing2_MainSave
    Datas() As Muing2_Save
End Type
Dim MySave As Muing2_MainSave
Function RSave(ByVal Title As String) As String
    For i = 0 To UBound(MySave.Datas)
        If MySave.Datas(i).Title = Title Then RSave = MySave.Datas(i).value: Exit For
    Next
End Function
Sub WSave(ByVal Title As String, ByVal value As String)
    For i = 0 To UBound(MySave.Datas)
        If MySave.Datas(i).Title = Title Then
            MySave.Datas(i).value = value
            GoTo Record
        End If
    Next
    
    ReDim Preserve MySave.Datas(UBound(MySave.Datas) + 1)
    MySave.Datas(UBound(MySave.Datas)).Title = Title
    MySave.Datas(UBound(MySave.Datas)).value = value
    
Record:
    CreateFolder "D:\Redstone\Muing2\"
    Open "D:\Redstone\Muing2\save" For Binary As #1
    Put #1, , MySave
    Close #1
End Sub
Sub Init()
    ReDim MySave.Datas(0)
    If Dir("D:\Redstone\Muing2\save") <> "" Then
        Open "D:\Redstone\Muing2\save" For Binary As #1
        Get #1, , MySave
        Close #1
        For i = 1 To UBound(Settings)
            Settings(i) = Val(RSave("Settings" & i))
        Next
    End If
End Sub
