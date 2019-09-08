VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Images"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BlendFunction As Long) ' As Long

Private Type BlendFunction
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public CompatibleDC As Long
Public Graphics As Long, name As String
Private DIB As Long
Private mWidth As Long, mHeight As Long
Private Blend As BlendFunction

Public Property Get Height() As Long 'DC高度
    Height = mHeight
End Property

Public Property Get Width() As Long 'DC宽度
    Width = mWidth
End Property

Public Sub CreateFromPath(hdc As Long, ByVal Path As String, Optional ByVal Scales As Single = 1)
    Dim Image As Long, W As Long, H As Long
    CreateImage StrPtr(Path), Image
    GdipGetImageWidth Image, W
    GdipGetImageHeight Image, H
    Call Create(hdc, W * Scales, H * Scales)
    GdipDrawImageRect Graphics, Image, 0, 0, Int(W * Scales), Int(H * Scales)
    DelImage Image
End Sub

Public Sub Create(hdc As Long, Width As Long, Height As Long)
    Dim bmHeader As BITMAPINFOHEADER

    mWidth = Width
    mHeight = Height

    With bmHeader
        .biBitCount = 32
        .biHeight = Height
        .biWidth = Width
        .biPlanes = 1
        .biSizeImage = (.biWidth * .biBitCount + 31) / 32 * 4 * .biHeight
        .biSize = Len(bmHeader)
    End With

    With Blend
        .AlphaFormat = &H1
        .BlendOp = 0
        .BlendFlags = &H0
    End With
    
    CompatibleDC = CreateCompatibleDC(hdc)
    DIB = CreateDIBSection(CompatibleDC, bmHeader, DIB_RGB_COLORS, ByVal 0, 0, 0)
    DeleteObject SelectObject(CompatibleDC, DIB)

    GdipCreateFromHDC CompatibleDC, Graphics
    '画图质量增强
    GdipSetSmoothingMode Graphics, SmoothingModeAntiAlias '抗锯齿
    GdipSetTextRenderingHint Graphics, TextRenderingHintAntiAlias '柔和字体
    
    AddMember Me
End Sub

Public Sub PresentWithClip(ByVal DstDC As Long, x As Long, y As Long, offsetX As Long, offsetY As Long, Width As Long, Height As Long, Optional Alpha As Long = 255)
    Dim BlendFunLong As Long

    If Alpha > 255 Then Alpha = 255
    If Width > mWidth Then Width = mWidth
    
    Blend.SourceConstantAlpha = Alpha
    CopyMemory BlendFunLong, Blend, 4
    
    AlphaBlend DstDC, x, y, Width, Height, CompatibleDC, offsetX, offsetY, Width, Height, BlendFunLong

    LastDrawC = LastDrawC + 1
End Sub

Public Sub Present(ByVal DstDC As Long, ByVal x As Long, ByVal y As Long, Optional Alpha As Long = 255)
    Dim BlendFunLong As Long

    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    
    Blend.SourceConstantAlpha = Alpha
    CopyMemory BlendFunLong, Blend, 4

    AlphaBlend DstDC, x, y, mWidth, mHeight, CompatibleDC, 0, 0, mWidth, mHeight, BlendFunLong
    
    LastDrawC = LastDrawC + 1
End Sub

Public Function PresentWithCtrl(ByVal DstDC As Long, Graphics As Long, ByVal x As Long, ByVal y As Long, value As Variant, ForeColor As Long, FontSize As Long, FontStyle As FontModes, Ctrl As CtrlClass, Optional Alpha As Long = 255, Optional Value2 As Variant)
    CtrlX = x: CtrlY = y
    Select Case Ctrl
        Case CtrlClass.Button
            CtrlW = mWidth: CtrlH = mHeight
            Call Present(DstDC, x, y, IIf(IsMouseIn, 255, 220))
            GameFont.DrawText Graphics, x + mWidth / 2, y + mHeight / 2 - FontSize / 0.75 / 2, 0, 0, value, ForeColor, center, FontSize, FontStyle
        Case CtrlClass.Button2
            CtrlW = mWidth: CtrlH = mHeight / 2 - 1
            Call PresentWithClip(DstDC, x, y, 0, IIf(IsMouseIn, Int(mHeight / 2), 0), mWidth, Int(mHeight / 2), 255)
            GameFont.DrawText Graphics, x + mWidth / 2, y + mHeight / 4 - FontSize / 0.75 / 2, 0, 0, value, ForeColor, center, FontSize, FontStyle
        Case CtrlClass.CheckBox
            CtrlW = mWidth: CtrlH = Int(mHeight / 3)
            If Value2 = 0 Then
                Call PresentWithClip(DstDC, x, y, 0, IIf(IsMouseIn, CtrlH, 0), mWidth, CtrlH, 255)
            Else
                Call PresentWithClip(DstDC, x, y, 0, CtrlH * 2, mWidth, CtrlH, 255)
            End If
            GameFont.DrawText Graphics, x + mWidth + 10, y + CtrlH / 2 - FontSize / 0.75 / 2, 0, 0, value, ForeColor, near, FontSize, FontStyle
            If IsMouseUp = True Then Value2 = IIf(Value2 = -1, 0, -1)
        Case CtrlClass.ProgressBar
            CtrlW = mWidth: CtrlH = mHeight / 2
            Call PresentWithClip(DstDC, x, y, 0, 0, mWidth, CtrlH, 255)
            Call PresentWithClip(DstDC, x, y, 0, CtrlH, mWidth * (value / Value2), CtrlH, 255)
        Case CtrlClass.SliderBar
            CtrlW = mWidth: CtrlH = mHeight / 2
            Call PresentWithClip(DstDC, x, y, 0, 0, mWidth, CtrlH, 255)
            Call PresentWithClip(DstDC, x, y, 0, CtrlH, mWidth * (value / Value2), CtrlH, 255)
            If IsMouseDown = True Then value = (MouseX - x) / mWidth * Value2
        Case CtrlClass.VScrollBar
            CtrlW = mWidth: CtrlH = mHeight - mWidth * 2
            Call PresentWithClip(DstDC, x, y, 0, mWidth, mWidth, mHeight - mWidth, 255)
            Call PresentWithClip(DstDC, x, y + (value / Value2) * (mHeight - mWidth * 2), 0, 0, mWidth, mWidth, 255)
            If IsMouseDown = True Then value = ((MouseY - mWidth / 2) - y) / (mHeight - mWidth * 2) * Value2
            If value < 0 Then value = 0
            If value > Value2 Then value = Value2
        Case CtrlClass.VScrollBar2
            CtrlW = mWidth: CtrlH = 360
            Call Present(DstDC, x, y + value / Value2 * 360, 255)
            If IsMouseDown = True Then value = ((MouseY - mHeight / 2) - y) / (360) * Value2
            If value < 0 Then value = 0
            If value > Value2 Then value = Value2
        Case CtrlClass.HScrollBar
            CtrlW = mWidth - mHeight * 2: CtrlH = mHeight
            Call PresentWithClip(DstDC, x, y, mHeight, 0, mWidth - mHeight, mHeight, 255)
            Call PresentWithClip(DstDC, x + (value / Value2) * (mWidth - mHeight * 2), y, 0, 0, mHeight, mHeight, 255)
            If IsMouseDown = True Then
                value = ((MouseX - mHeight / 2) - x) / (mWidth - mHeight * 2) * Value2
            End If
            If value < 0 Then value = 0
            If value > Value2 Then value = Value2
        Case CtrlClass.EditBox
            CtrlW = mWidth: CtrlH = mHeight / 2
            Call PresentWithClip(DstDC, x, y, 0, IIf(EditBoxs(value), Int(mHeight / 2), 0), mWidth, Int(mHeight / 2), 255)
            If Not IsMissing(Value2) Then
                Dim FFF As String
                For i = 1 To Len(EditBoxText(value))
                    FFF = FFF & "*"
                Next
                GameFont.DrawText Graphics, x + Int(mHeight / 4), y + (Int(mHeight / 2) / 2 - FontSize / 0.75 / 2), mWidth - Int(mHeight / 2), 0, FFF, ForeColor, near, FontSize, FontStyle
            Else
                GameFont.DrawText Graphics, x + Int(mHeight / 4), y + (Int(mHeight / 2) / 2 - FontSize / 0.75 / 2), mWidth - Int(mHeight / 2), 0, EditBoxText(value), ForeColor, near, FontSize, FontStyle
            End If
            If IsClick Then EditBoxs(LastBox) = False: EditBoxs(value) = True: LastBox = value
            If EditBoxs(value) Then
                NowEdit = value
                If MouseState <> 0 And (Not IsClick) Then EditBoxs(value) = False
            End If
            If Len(EditBoxText(value)) > 20 Then EditBoxText(value) = Left(EditBoxText(value), 20)
    End Select
End Function
Sub SetClickArea(ByVal x As Long, ByVal y As Long)
    CtrlX = x: CtrlY = y
    CtrlW = mWidth: CtrlH = mHeight
End Sub
Sub Dispose()
    GdipDeleteGraphics Graphics
    DeleteObject CompatibleDC
    DeleteObject DIB
End Sub
