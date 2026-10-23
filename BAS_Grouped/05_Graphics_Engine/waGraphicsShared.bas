' =============================================================================
' Group: 05_Graphics_Engine
' Source module: waGraphicsShared.bas
' Procedures: 7
' =============================================================================

' -----------------------------------------------------------------------------
' CreatePen (Function), source lines 55-58
' -----------------------------------------------------------------------------
Public Function CreatePen(Optional ColorRGB As Long = 6579300, Optional Opacity As Long = 100, Optional Width As Single = 1) As waPen
    Set CreatePen = New waPen
    CreatePen.SetColor ColorRGB, Opacity, Width
End Function

' -----------------------------------------------------------------------------
' CreateBrush (Function), source lines 59-62
' -----------------------------------------------------------------------------
Public Function CreateBrush(Optional ColorRGB As Long = 6579300, Optional Opacity As Long = 100) As waBrush
    Set CreateBrush = New waBrush
    CreateBrush.SetColor ColorRGB, Opacity
End Function

' -----------------------------------------------------------------------------
' ColorFromOle (Function), source lines 63-67
' -----------------------------------------------------------------------------
Public Function ColorFromOle(ColorOle As Long) As Long
    Dim c As Long
    c = ColorOle And &HFFFFFF
    ColorFromOle = GetSysColor(c)
End Function

' -----------------------------------------------------------------------------
' BrightnessRGBA (Function), source lines 68-72
' -----------------------------------------------------------------------------
Public Function BrightnessRGBA(rgba_ As Long) As Byte
    Dim tmprgb(3) As Byte
    CopyMemory tmprgb(0), rgba_, 4
    BrightnessRGBA = CByte(0.2126 * tmprgb(2) + 0.7152 * tmprgb(1) + 0.0722 * tmprgb(0))
End Function

' -----------------------------------------------------------------------------
' BrightnessRGB (Function), source lines 73-75
' -----------------------------------------------------------------------------
Public Function BrightnessRGB(R As Long, g As Long, b As Long) As Byte
    BrightnessRGB = CByte(0.2126 * R + 0.7152 * g + 0.0722 * b)
End Function

' -----------------------------------------------------------------------------
' BrightnessRGBLong (Function), source lines 76-82
' -----------------------------------------------------------------------------
Public Function BrightnessRGBLong(RGB_ As Long) As Byte
    Dim R As Long, g As Long, b As Long
    b = RGB_ \ 65536
    g = (RGB_ - b * 65536) \ 256
    R = RGB_ - b * 65536 - g * 256
    BrightnessRGBLong = CByte(0.2126 * R + 0.7152 * g + 0.0722 * b)
End Function

' -----------------------------------------------------------------------------
' ConvertColor (Function), source lines 83-90
' -----------------------------------------------------------------------------
Public Function ConvertColor(Color As Long, Opacity As Long) As Long
    Dim BGRA(3) As Byte
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function
