' =============================================================================
' Group: 05_Graphics_Engine
' Source module: waPen.cls
' Procedures: 11
' =============================================================================

' -----------------------------------------------------------------------------
' matrix (PropertySet), source lines 54-56
' -----------------------------------------------------------------------------
Property Set matrix(oValue As waMatrix)
    GdipSetPenTransform lPtr, oValue.Ptr
End Property

' -----------------------------------------------------------------------------
' EndCap (PropertyLet), source lines 57-59
' -----------------------------------------------------------------------------
Property Let EndCap(oValue As LineCap)
    GdipSetPenEndCap lPtr, oValue
End Property

' -----------------------------------------------------------------------------
' EndCap (PropertyGet), source lines 60-62
' -----------------------------------------------------------------------------
Property Get EndCap() As LineCap
    GdipGetPenEndCap lPtr, EndCap
End Property

' -----------------------------------------------------------------------------
' DashStyle (PropertyLet), source lines 63-65
' -----------------------------------------------------------------------------
Property Let DashStyle(oValue As DashStyle)
    GdipSetPenDashStyle lPtr, oValue
End Property

' -----------------------------------------------------------------------------
' DashStyle (PropertyGet), source lines 66-68
' -----------------------------------------------------------------------------
Property Get DashStyle() As DashStyle
    GdipGetPenDashStyle lPtr, DashStyle
End Property

' -----------------------------------------------------------------------------
' Ptr (PropertyGet), source lines 69-72
' -----------------------------------------------------------------------------
Property Get Ptr() As LongPtr
    If Not Initialized Then SetColor DefColor, 100
    Ptr = lPtr
End Property

' -----------------------------------------------------------------------------
' ColorARGB (PropertyLet), source lines 73-75
' -----------------------------------------------------------------------------
Property Let ColorARGB(oValue As Long)
    GdipSetPenColor lPtr, oValue
End Property

' -----------------------------------------------------------------------------
' ColorARGB (PropertyGet), source lines 76-78
' -----------------------------------------------------------------------------
Property Get ColorARGB() As Long
    GdipGetPenColor lPtr, ColorARGB
End Property

' -----------------------------------------------------------------------------
' SetColor (Sub), source lines 80-84
' -----------------------------------------------------------------------------
Public Sub SetColor(ColorRGB As Long, Opacity As Long, Optional Width As Single = 1)
    If Initialized Then Class_Terminate
    GdipCreatePen1 ConvertColor(ColorRGB, Opacity), Width, 2, lPtr
    Initialized = True
End Sub

' -----------------------------------------------------------------------------
' ConvertColor (Function), source lines 86-93
' -----------------------------------------------------------------------------
Private Function ConvertColor(Color As Long, Opacity As Long) As Long
    Dim BGRA(3) As Byte
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 94-96
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    If Initialized Then GdipDeletePen lPtr
End Sub
