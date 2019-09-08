VERSION 5.00
Begin VB.Form SongCreateWindow 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "谱子"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   ControlBox      =   0   'False
   Icon            =   "SongCreateWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   547
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox IDText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   225
      Left            =   4950
      MaxLength       =   5
      TabIndex        =   17
      Text            =   "0"
      Top             =   1350
      Width           =   2565
   End
   Begin VB.Timer Checker 
      Interval        =   100
      Left            =   7500
      Top             =   150
   End
   Begin VB.TextBox HardText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009002FF&
      Height          =   225
      Left            =   6300
      TabIndex        =   14
      Text            =   "1.5"
      Top             =   2550
      Width           =   1215
   End
   Begin VB.TextBox NormalText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0B000&
      Height          =   285
      Left            =   6300
      TabIndex        =   12
      Text            =   "1"
      Top             =   2250
      Width           =   1215
   End
   Begin VB.TextBox EasyText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009FD5&
      Height          =   225
      Left            =   6300
      TabIndex        =   10
      Text            =   "0.9"
      Top             =   1950
      Width           =   1215
   End
   Begin VB.TextBox DownloadText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   3150
      Width           =   6915
   End
   Begin VB.TextBox ArtistText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   225
      Left            =   600
      TabIndex        =   5
      Top             =   2550
      Width           =   4065
   End
   Begin VB.TextBox MusicText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   1950
      Width           =   4065
   End
   Begin VB.TextBox MakerText 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   1350
      Width           =   4065
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "谱子设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007C7976&
      Height          =   285
      Left            =   150
      TabIndex        =   18
      Top             =   150
      Width           =   780
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "谱子ID"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   4950
      TabIndex        =   16
      Top             =   1050
      Width           =   600
   End
   Begin VB.Label OKBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E8E8E8&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   315
      Left            =   6750
      TabIndex        =   15
      Top             =   3750
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008D02FF&
      Height          =   285
      Left            =   4950
      TabIndex        =   13
      Top             =   2550
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0B000&
      Height          =   285
      Left            =   4950
      TabIndex        =   11
      Top             =   2250
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009FD5&
      Height          =   285
      Left            =   4950
      TabIndex        =   9
      Top             =   1950
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "谱子对象移动速度(*)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   4950
      TabIndex        =   8
      Top             =   1650
      Width           =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "下载地址（百度云盘地址，上传时必须输入）"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00919191&
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   2850
      Width           =   3900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "音乐作曲者(*)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   2250
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "音乐名称(*)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "谱子制作者(*)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BAB539&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1050
      Width           =   1185
   End
End
Attribute VB_Name = "SongCreateWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oShadow As New aShadow

Private Sub Checker_Timer()
    Check
End Sub

Private Sub Form_Load()

    With oShadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 20
            .Transparency = 8
        End If
    End With

    ArtistText.Text = EditMap.Artist
    MusicText.Text = EditMap.Title
    MakerText.Text = EditMap.Maker
    DownloadText.Text = EditMap.DownloadSource
    EasyText.Text = EditMap.Levels(0).ObjSpeed
    NormalText.Text = EditMap.Levels(1).ObjSpeed
    HardText.Text = EditMap.Levels(2).ObjSpeed
    IDText.Text = EditMap.MapID
    MainWindow.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainWindow.Enabled = True
    MainWindow.SetFocus
    
    Set oShadow = Nothing
End Sub

Sub Check()
    OKBtn.Visible = True
    If MusicText.Text = "" Then OKBtn.Visible = False
    Label2.ForeColor = IIf(MusicText.Text = "", RGB(255, 0, 0), RGB(51, 181, 186))
    If MakerText.Text = "" Then OKBtn.Visible = False
    Label1.ForeColor = IIf(MakerText.Text = "", RGB(255, 0, 0), RGB(51, 181, 186))
    If HardText.Text = "" Then OKBtn.Visible = False
    HardText.BackColor = IIf(HardText.Text = "", RGB(255, 197, 197), RGB(249, 249, 249))
    If NormalText.Text = "" Then OKBtn.Visible = False
    NormalText.BackColor = IIf(NormalText.Text = "", RGB(255, 197, 197), RGB(249, 249, 249))
    If EasyText.Text = "" Then OKBtn.Visible = False
    EasyText.BackColor = IIf(EasyText.Text = "", RGB(255, 197, 197), RGB(249, 249, 249))
    If ArtistText.Text = "" Then OKBtn.Visible = False
    Label3.ForeColor = IIf(ArtistText.Text = "", RGB(255, 0, 0), RGB(51, 181, 186))
End Sub

Private Sub IDText_Change()
    IDText = Int(Val(IDText.Text))
End Sub

Private Sub OKBtn_Click()
    If Val(IDText.Text) = 0 Then MsgBox "谱子ID不能为0！", 16, "谱子编辑器": Exit Sub

    If IDText.Enabled = True Then
    
        If Connected = False Then MsgBox "很抱歉，我们需要连接服务器来检查你的谱子ID是否被占用。", 48, "谱子编辑器": Exit Sub
        If Newest = False Then MsgBox "很抱歉，你的游戏版本不是最新的。", 48, "谱子编辑器": Exit Sub
        
        MapCheck = ""
        Send "checkmap*" & ArtistText.Text & "*" & MakerText.Text & "*" & MusicText.Text & "*" & IDText.Text
        
        Do While MapCheck = ""
            Sleep 10: DoEvents
        Loop
        Dim temp() As String
        temp = Split(MapCheck, "*")
        If temp(1) <> "yes" Then MsgBox "该谱子ID被占用！", 48, "谱子编辑器": Exit Sub
        
    End If
    
    If CheckText(ArtistText.Text) Then MsgBox "作曲者名称存在不允许的字符。", 16, "谱子编辑器": Exit Sub
    If CheckText(MusicText.Text) Then MsgBox "音乐名称存在不允许的字符。", 16, "谱子编辑器": Exit Sub
    If CheckText(MakerText.Text) Then MsgBox "谱子作者名称存在不允许的字符。", 16, "谱子编辑器": Exit Sub
    
    EditMap.Artist = ArtistText.Text: EditMap.Title = MusicText.Text: EditMap.Maker = MakerText.Text
    EditMap.Levels(0).ObjSpeed = Val(EasyText.Text): EditMap.Levels(1).ObjSpeed = Val(NormalText.Text): EditMap.Levels(2).ObjSpeed = Val(HardText.Text)
    EditMap.DownloadSource = DownloadText.Text
    EditMap.MapID = Val(IDText.Text)
    If EditMap.DownloadSource = "" Then EditMap.DownloadSource = "/"
    If Dir(SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\", vbDirectory) = "" Then MkDir SongPath & "\" & EditMap.MapID & " - " & EditMap.Title & " - " & EditMap.Maker & "\"
    
    Unload Me
End Sub
