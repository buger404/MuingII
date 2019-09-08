VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SetPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Draw As New Images
Dim ScrollY As Long
Public Property Get Visible() As Boolean
    Visible = (mNowShow = "SetPage")
End Property

Public Sub Display()
    If DebugMode = False Then On Error Resume Next
    
    NowPage = "SetPage"
    
    Call DrawPage0
    
    Draw.Present GameDraw.CompatibleDC, 0, 0
End Sub
Public Sub DrawPage0()
    BackImg.ImageByIndex(6).Present Draw.CompatibleDC, 0, 0, 255

    MainImg.Image("passframe.png").Present Draw.CompatibleDC, 57, 130
    
    If GameSave.RSave("UserID") = "" Then
        GameFont.DrawText Draw.Graphics, 127, 147, GWW, 30, "����û�д����κ�����", argb(255, 27, 27, 27), near, 16, Regular
    Else
        GameFont.DrawText Draw.Graphics, 127, 147, GWW, 30, "�Ѵ��� ID " & GameSave.RSave("UserID") & " ������", argb(255, 27, 27, 27), near, 16, Regular
        With EditImg.Image("erase.png")
            .SetClickArea GWW - 110, 137
            .Present Draw.CompatibleDC, GWW - 110, 137, IIf(IsMouseIn, 255, 200)
            If IsClick Then
                If GameNotify.Asks(MuingAsk, "�˺󲻻��ٽ����Զ���½�������������������������", "ȡ��", "����") = 0 Then Exit Sub
                GameSave.WSave "UserID", ""
                GameSave.WSave "Password", ""
            End If
        End With
    End If
    
    GameFont.DrawText Draw.Graphics, 137, 247, GWW, 30, "ɾ����������������", argb(255, 255, 0, 0), near, 16, Regular
    With EditImg.Image("erase.png")
        .SetClickArea 72, 237
        .Present Draw.CompatibleDC, 72, 237, IIf(IsMouseIn, 255, 200)
        If IsClick Then
            If GameNotify.Asks(MuingAsk, "ɾ���������е����ӣ��˲��������棡����", "ȡ��", "ɾ��") = 0 Then Exit Sub
            Dim FSO As Object
            Set FSO = CreateObject("Scripting.FileSystemObject")
            On Error Resume Next
            GameBGM.StopMusic
            GameBGM.Dispose
            FSO.DeleteFolder SongPath
            Err.Clear
            GetSongList
        End If
    End With
    
    Dim s As Integer
    Dim DesList() As String
    ReDim DesList(UBound(Settings))
    
    DesList(MuingII_Settings.MuingBlurBack) = "ȡ������ģ��Ч��"
    DesList(MuingII_Settings.MuingNoServer) = "����ģʽ"
    DesList(MuingII_Settings.MuingShowHot) = "ȡ����ʾ��̬Ƶ��"
    DesList(MuingII_Settings.MuingShowHot2) = "ȡ����̬����"
    DesList(MuingII_Settings.MuingUseDog) = "Autoģʽ�²�Ҫʹ�ù�����"
    DesList(MuingII_Settings.MuingDebug) = "��ʾ������Ϣ"
    DesList(MuingII_Settings.MuingShowAllLines) = "��Ϸʱ��ʾ���е���"
    
    For s = 1 To UBound(Settings)
        GameFont.DrawText Draw.Graphics, 127, 307 + 5 + 40 * (s - 1), GWW, 30, DesList(s), argb(255, 255, 255, 255), near, 16, Regular
        With MainImg.Image("check" & Settings(s) & ".png")
            .SetClickArea 80, 307 + 40 * (s - 1)
            .Present Draw.CompatibleDC, 80, 307 + 40 * (s - 1), IIf(IsMouseIn, 255, 200)
            If IsClick Then
                ChangeSettings s, IIf(Settings(s) = 0, 1, 0)
            End If
        End With
    Next
End Sub

Private Sub Class_Initialize()
    Draw.Create TargetDC, GWW, GWH
End Sub



