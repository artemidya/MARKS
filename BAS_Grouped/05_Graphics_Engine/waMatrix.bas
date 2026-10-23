' =============================================================================
' Group: 05_Graphics_Engine
' Source module: waMatrix.cls
' Procedures: 6
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 60-62
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    GdipCreateMatrix lPtr
End Sub

' -----------------------------------------------------------------------------
' Ptr (PropertyGet), source lines 63-65
' -----------------------------------------------------------------------------
Property Get Ptr() As LongPtr
    Ptr = lPtr
End Property

' -----------------------------------------------------------------------------
' Translate (Sub), source lines 66-68
' -----------------------------------------------------------------------------
Public Sub Translate(offsetX As Long, offsetY As Long)
    GdipTranslateMatrix lPtr, CSng(offsetX), CSng(offsetY), MatrixOrderAppend
End Sub

' -----------------------------------------------------------------------------
' ScaleM (Sub), source lines 69-71
' -----------------------------------------------------------------------------
Public Sub ScaleM(scaleX As Single, scaleY As Single)
    GdipScaleMatrix lPtr, scaleX, scaleY, MatrixOrderAppend
End Sub

' -----------------------------------------------------------------------------
' Rotate (Sub), source lines 72-74
' -----------------------------------------------------------------------------
Public Sub Rotate(Angel As Single)
    GdipRotateMatrix lPtr, Angel, MatrixOrderAppend
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 77-79
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    GdipDeleteMatrix lPtr
End Sub
