' =============================================================================
' Group: 05_Graphics_Engine
' Source module: waGraphics.cls
' Procedures: 13
' =============================================================================

' -----------------------------------------------------------------------------
' Init (Sub), source lines 80-90
' -----------------------------------------------------------------------------
Private Sub Init()
'    lhDC = GetDC(lHWndForm)
    Dim gdipStartupInput As GDIPlusStartupInput
    gdipStartupInput.GdiPlusVersion = 1&
    GdiplusStartup lGdipToken, gdipStartupInput, ByVal 0
    GdipCreateFromHDC lhDC, lHGraphics
    GdipSetSmoothingMode lHGraphics, SmoothingModeAntiAlias
'    Matrix.ScaleM 1, -1

    Initialized = True
End Sub

' -----------------------------------------------------------------------------
' hDC (PropertyLet), source lines 91-100
' -----------------------------------------------------------------------------
Property Let hDC(oValue As LongPtr)
1
    If Not Initialized Then
        lhDC = oValue
        Init
    Else
        Class_Terminate
        GoTo 1
    End If
End Property

' -----------------------------------------------------------------------------
' hDC (PropertyGet), source lines 101-103
' -----------------------------------------------------------------------------
Property Get hDC() As LongPtr
    hDC = lhDC
End Property

' -----------------------------------------------------------------------------
' DrawLine (Sub), source lines 105-107
' -----------------------------------------------------------------------------
Public Sub DrawLine(pen As waPen, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    GdipDrawLineI lHGraphics, pen.Ptr, x1, y1, x2, y2
End Sub

' -----------------------------------------------------------------------------
' DrawBezier (Sub), source lines 108-110
' -----------------------------------------------------------------------------
Public Sub DrawBezier(pen As waPen, x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long, x4 As Long, y4 As Long)
    GdipDrawBezierI lHGraphics, pen.Ptr, x1, y1, x2, y2, x3, y3, x4, y4
End Sub

' -----------------------------------------------------------------------------
' DrawRect (Sub), source lines 111-113
' -----------------------------------------------------------------------------
Public Sub DrawRect(pen As waPen, x As Long, y As Long, Width As Long, Height As Long)
    GdipDrawRectangleI lHGraphics, pen.Ptr, x, y, Width, Height
End Sub

' -----------------------------------------------------------------------------
' FillRect (Sub), source lines 114-116
' -----------------------------------------------------------------------------
Public Sub FillRect(brush As waBrush, x As Long, y As Long, Width As Long, Height As Long)
    GdipFillRectangleI lHGraphics, brush.Ptr, x, y, Width, Height
End Sub

' -----------------------------------------------------------------------------
' SetWorldTransform (Sub), source lines 117-119
' -----------------------------------------------------------------------------
Public Sub SetWorldTransform(matrix As waMatrix)
    GdipSetWorldTransform lHGraphics, matrix.Ptr
End Sub

' -----------------------------------------------------------------------------
' ResetWorldTransform (Sub), source lines 120-122
' -----------------------------------------------------------------------------
Public Sub ResetWorldTransform()
    GdipResetWorldTransform lHGraphics
End Sub

' -----------------------------------------------------------------------------
' SetClip (Sub), source lines 123-125
' -----------------------------------------------------------------------------
Public Sub SetClip(x As Long, y As Long, Width As Long, Height As Long, Optional CombineMode As CombineMode = CombineModeReplace)
    GdipSetClipRectI lHGraphics, x, y, Width, Height, CombineMode
End Sub

' -----------------------------------------------------------------------------
' ResetClip (Sub), source lines 126-128
' -----------------------------------------------------------------------------
Public Sub ResetClip()
    GdipResetClip lHGraphics
End Sub

' -----------------------------------------------------------------------------
' Render (Sub), source lines 129-157
' -----------------------------------------------------------------------------
Public Sub Render(Optional Draw As Boolean = True, Optional Fill As Boolean = False, Optional ShadowBrush As waBrush, Optional SpecialPen As waPen)
    If Not Curves Is Nothing Then
        Dim curve As waPath
        If ShadowBrush Is Nothing Then
            If SpecialPen Is Nothing Then
                For Each curve In Curves
                    With curve
                        If Fill Then GdipFillPath lHGraphics, .brush.Ptr, .Ptr
                        If Draw Then GdipDrawPath lHGraphics, .pen.Ptr, .Ptr
                    End With
                Next
            Else
                For Each curve In Curves
                    With curve
                        If Fill Then GdipFillPath lHGraphics, .brush.Ptr, .Ptr
                        If Draw Then GdipDrawPath lHGraphics, SpecialPen.Ptr, .Ptr
                    End With
                Next
            End If
        Else
            'Shadow
            For Each curve In Curves
                With curve
                    GdipFillPath lHGraphics, ShadowBrush.Ptr, .Ptr
                End With
            Next
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 159-166
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    If Initialized Then
        GdipDeleteGraphics lHGraphics
        GdiplusShutdown lGdipToken
'        ReleaseDC lHWndForm, lhDC
        Initialized = False
    End If
End Sub
