' =============================================================================
' Group: 05_Graphics_Engine
' Source module: waBrush.cls
' Procedures: 12
' =============================================================================

' -----------------------------------------------------------------------------
' Brightness (PropertyGet), source lines 40-46
' -----------------------------------------------------------------------------
Property Get Brightness() As Byte
    Dim tmprgb(3) As Byte
    Dim Color As Long
    GdipGetSolidFillColor lPtr, Color
    CopyMemory tmprgb(0), Color, 4
    Brightness = CByte(0.2126 * tmprgb(2) + 0.7152 * tmprgb(1) + 0.0722 * tmprgb(0))
End Property

' -----------------------------------------------------------------------------
' Color (PropertyGet), source lines 47-60
' -----------------------------------------------------------------------------
Property Get Color() As RGBA
    Dim tmprgb(3) As Byte
    Dim colorLong As Long

    GdipGetSolidFillColor lPtr, colorLong
    CopyMemory tmprgb(0), colorLong, 4
    With Color
        .R = tmprgb(2)
        .g = tmprgb(1)
        .b = tmprgb(0)
        .a = tmprgb(3)
    End With
    
End Property

' -----------------------------------------------------------------------------
' P1_GradientLine (PropertyGet), source lines 61-63
' -----------------------------------------------------------------------------
Property Get P1_GradientLine() As POINTL
    P1_GradientLine = P1
End Property

' -----------------------------------------------------------------------------
' P2_GradientLine (PropertyGet), source lines 64-66
' -----------------------------------------------------------------------------
Property Get P2_GradientLine() As POINTL
    P2_GradientLine = P2
End Property

' -----------------------------------------------------------------------------
' ColorLineGradient1 (PropertyGet), source lines 68-80
' -----------------------------------------------------------------------------
Property Get ColorLineGradient1() As RGBA
    Dim tmprgb(3) As Byte
    Dim colorLongPtr As LongPtr
    GdipGetLineColors lPtr, colorLongPtr
    CopyMemory tmprgb(0), colorLongPtr, 4
    
    With ColorLineGradient1
        .R = tmprgb(2)
        .g = tmprgb(1)
        .b = tmprgb(0)
        .a = tmprgb(3)
    End With
End Property

' -----------------------------------------------------------------------------
' ColorLineGradient2 (PropertyGet), source lines 81-93
' -----------------------------------------------------------------------------
Property Get ColorLineGradient2() As RGBA
    Dim tmprgb(7) As Byte
    Dim colorLongPtr As LongPtr
    GdipGetLineColors lPtr, colorLongPtr
    CopyMemory tmprgb(0), colorLongPtr, 8

    With ColorLineGradient2
        .R = tmprgb(6)
        .g = tmprgb(5)
        .b = tmprgb(4)
        .a = tmprgb(7)
    End With
End Property

' -----------------------------------------------------------------------------
' Opacity (PropertyGet), source lines 95-101
' -----------------------------------------------------------------------------
Property Get Opacity() As Byte
    Dim tmprgb(3) As Byte
    Dim Color As Long
    GdipGetSolidFillColor lPtr, Color
    CopyMemory tmprgb(0), Color, 4
    Opacity = tmprgb(3)
End Property

' -----------------------------------------------------------------------------
' Ptr (PropertyGet), source lines 102-105
' -----------------------------------------------------------------------------
Property Get Ptr() As LongPtr
    If Not Initialized Then SetColor DefColor, 100
    Ptr = lPtr
End Property

' -----------------------------------------------------------------------------
' FillType (PropertyGet), source lines 106-108
' -----------------------------------------------------------------------------
Property Get FillType() As BrushType
    GdipGetBrushType lPtr, FillType
End Property

' -----------------------------------------------------------------------------
' SetColor (Sub), source lines 109-113
' -----------------------------------------------------------------------------
Public Sub SetColor(ColorRGB As Long, Opacity As Long)
    If Initialized Then Class_Terminate
    GdipCreateSolidFill ConvertColor(ColorRGB, Opacity), lPtr
    Initialized = True
End Sub

' -----------------------------------------------------------------------------
' SetLinearGradient (Sub), source lines 119-129
' -----------------------------------------------------------------------------
Public Sub SetLinearGradient(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ColorRGB1 As Long, ColorRGB2 As Long, Optional WrapMode As WrapMode = WrapModeTileFlipXY, Optional Opacity1 As Long = 100, Optional Opacity2 As Long = 100)
    If Initialized Then Class_Terminate
    
    P1.x = x1
    P1.y = y1
    P2.x = x2
    P2.y = y2
    
    GdipCreateLineBrushI P1, P2, ConvertColor(ColorRGB1, Opacity1), ConvertColor(ColorRGB2, Opacity2), WrapMode, lPtr
    Initialized = True
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 131-133
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    If Initialized Then GdipDeleteBrush lPtr
End Sub
