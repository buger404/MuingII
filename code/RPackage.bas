Attribute VB_Name = "RPackage"
Dim MyFileList As RFileList
Private Type RPackage
    data() As Byte
    Path As String
End Type
Private Type RFileList
    Files() As RPackage
End Type
Sub Export(ByVal Package As String, ByVal Path As String)
    Open Package For Binary As #1
    Get #1, , MyFileList
    Close #1
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    For i = 0 To UBound(MyFileList.Files)
        CreateFolder Path & MyFileList.Files(i).Path
        Open Path & MyFileList.Files(i).Path For Binary As #1
        Put #1, , MyFileList.Files(i).data
        Close #1
    If DebugMode = False Then On Error Resume Next
    Close #1
    Next
End Sub
Sub MakePackage(ByVal Package As String, ByVal Path As String)
    Dim temp() As String
    temp = DirAllFiles(Path)
    ReDim MyFileList.Files(UBound(temp) - 1)
    For i = 1 To UBound(temp)
        MyFileList.Files(i - 1).Path = Mid(temp(i), Len(Path) + 1)
        ReDim MyFileList.Files(i - 1).data(FileLen(temp(i)) - 1)
        Open temp(i) For Binary As #1
        Get #1, , MyFileList.Files(i - 1).data
        Close #1
    Next
    Open Package For Binary As #1
    Put #1, , MyFileList
    Close #1
End Sub
Function DirAllFiles(ByVal Path As String) As String()
    Dim DirTasks() As String, file As String, Folder As String
    Dim FileList() As String
    ReDim DirTasks(1), FileList(0)
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    DirTasks(1) = Path
    Do While UBound(DirTasks) > 0
        file = Dir(DirTasks(1))
        Do While file <> ""
            ReDim Preserve FileList(UBound(FileList) + 1)
            FileList(UBound(FileList)) = DirTasks(1) & file
            file = Dir()
            DoEvents
        Loop
        Folder = Dir(DirTasks(1), vbDirectory)
        Do While Folder <> ""
            If Folder <> "." And Folder <> ".." And (Not Folder Like "*.*") Then
                ReDim Preserve DirTasks(UBound(DirTasks) + 1)
                DirTasks(UBound(DirTasks)) = DirTasks(1) & Folder & "\"
                StateText.Caption = DirTasks(1) & Folder & "\"
            End If
            Folder = Dir(, vbDirectory)
            DoEvents
        Loop
        DirTasks(1) = DirTasks(UBound(DirTasks))
        ReDim Preserve DirTasks(UBound(DirTasks) - 1)
    Loop
    DirAllFiles = FileList
End Function



