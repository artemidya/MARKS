' =============================================================================
' Group: 05_Graphics_Engine
' Source module: waPath.cls
' Procedures: 25
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 44-51
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    'FillModeWinding
    GdipCreatePath FillModeAlternate, Ptr

    Set lPen = DefPen
    Set lBrush = DefBrush
    NeedRecalculateBounds = True
End Sub

' -----------------------------------------------------------------------------
' Reverse (Sub), source lines 52-54
' -----------------------------------------------------------------------------
Public Sub Reverse()
    GdipReversePath Ptr
End Sub

' -----------------------------------------------------------------------------
' GetPolyline (Function), source lines 55-68
' -----------------------------------------------------------------------------
Public Function GetPolyline(segment As segment, Optional ByVal flatness As Single = 1 / 4) As segment()
    Dim path As waPath
    Dim segments() As segment
    Set path = New waPath
    With segment
        path.AddBezierSegment .P1.x, .P1.y, .P2.x, .P2.y, .P3.x, .P3.y, .P4.x, .P4.y
        GdipFlattenPath path.Ptr, , flatness
        segments = path.segments
        segments(0).P2 = .P1
        segments(1).P1 = .P1
        segments(UBound(segments)).P2 = .P4
    End With
    GetPolyline = segments
End Function

' -----------------------------------------------------------------------------
' ChangeStartNode (Function), source lines 69-108
' -----------------------------------------------------------------------------
Public Function ChangeStartNode(NumberOfNewStartNode As Long) As waPath
    If NumberOfNewStartNode < 2 Then
        Set ChangeStartNode = Me
        Exit Function
    End If
    
    Dim segments() As segment
    segments = Me.segments
    
    Dim path As waPath
    Set path = New waPath
    path.StartSubPath
    With segments(NumberOfNewStartNode).P1
        path.AddLineSegment .x, .y
    End With
    Dim q As Long
    For q = NumberOfNewStartNode To UBound(segments)
        With segments(q)
            If .Type = CurveSegment Then
                path.AddBezierSegment .P1.x, .P1.y, .P2.x, .P2.y, .P3.x, .P3.y, .P4.x, .P4.y
            ElseIf .Type = LineSegment Then
                path.AddLineSegment .P2.x, .P2.y
            End If
        End With
    Next
    
    For q = 1 To NumberOfNewStartNode - 1
        With segments(q)
            If .Type = CurveSegment Then
                path.AddBezierSegment .P1.x, .P1.y, .P2.x, .P2.y, .P3.x, .P3.y, .P4.x, .P4.y
            ElseIf .Type = LineSegment Then
                path.AddLineSegment .P2.x, .P2.y
            End If
        End With
    Next
    Set path.brush = Me.brush
    Set path.pen = Me.pen
    Set ChangeStartNode = path
    
End Function

' -----------------------------------------------------------------------------
' closed (PropertyGet), source lines 109-121
' -----------------------------------------------------------------------------
Property Get closed() As Boolean
    Dim s() As segment
    s = Me.segments
    Dim n As Long
    n = UBound(s)
    With s(n)
        If .Type = CurveSegment Then
            If s(1).P1.x = .P4.x And s(1).P1.y = .P4.y Then closed = True
        Else
            If s(1).P1.x = .P2.x And s(1).P1.y = .P2.y Then closed = True
        End If
    End With
End Property

' -----------------------------------------------------------------------------
' segments (PropertyGet), source lines 122-192
' -----------------------------------------------------------------------------
Property Get segments() As segment()
    
    Dim count As Long
    GdipGetPointCount Ptr, count
    Dim p() As POINTL
    Dim t() As Byte
    
    ReDim p(count - 1)
    ReDim t(count - 1)
    GdipGetPathPointsI Ptr, p(0), count
    GdipGetPathTypes Ptr, t(0), count
    

    
    Dim cps As Long: cps = -1
    Dim q As Long
    Dim LastPoint As POINTL
    
    Dim countsegments As Long
    For q = 0 To count - 1
        Select Case t(q)
            Case 0  'Start
                countsegments = countsegments + 1
            Case 1  'Line
                countsegments = countsegments + 1
            Case 129 'Line (end)
                countsegments = countsegments + 1
            Case 3  'Bezier
                countsegments = countsegments + 1
                q = q + 2
        End Select
    Next
    
    Dim s() As segment
    ReDim s(countsegments - 1)
    
    For q = 0 To count - 1
        Select Case t(q)
            Case 0  'Start
                cps = cps + 1
                s(cps).P1 = LastPoint
                s(cps).P2 = p(q)
                s(cps).Type = InvisibleLine
                LastPoint = s(cps).P2
            Case 1  'Line
                cps = cps + 1
                s(cps).P1 = LastPoint
                s(cps).P2 = p(q)
                s(cps).Type = LineSegment
                LastPoint = s(cps).P2
            Case 129 'Line (end)
                cps = cps + 1
                s(cps).P1 = LastPoint
                s(cps).P2 = p(q)
                s(cps).Type = LineSegment
                LastPoint = s(cps).P2
            Case 3  'Bezier
                cps = cps + 1
                s(cps).P1 = LastPoint
                s(cps).P2 = p(q)
                s(cps).P3 = p(q + 1)
                s(cps).P4 = p(q + 2)
                LastPoint = s(cps).P4
                q = q + 2
                s(cps).Type = CurveSegment
                
        End Select
    Next

    segments = s
End Property

' -----------------------------------------------------------------------------
' LastSegment (PropertyGet), source lines 193-197
' -----------------------------------------------------------------------------
Property Get LastSegment() As segment
    Dim segments() As segment
    segments = Me.segments
    LastSegment = segments(UBound(segments))
End Property

' -----------------------------------------------------------------------------
' FirstSegment (PropertyGet), source lines 198-202
' -----------------------------------------------------------------------------
Property Get FirstSegment() As segment
    Dim segments() As segment
    segments = Me.segments
    FirstSegment = segments(1)
End Property

' -----------------------------------------------------------------------------
' FirstPoint (PropertyGet), source lines 203-207
' -----------------------------------------------------------------------------
Property Get FirstPoint() As POINTL
    Dim segments() As segment
    segments = Me.segments
    FirstPoint = segments(1).P1
End Property

' -----------------------------------------------------------------------------
' LastPoint (PropertyGet), source lines 208-218
' -----------------------------------------------------------------------------
Property Get LastPoint() As POINTL
    Dim segments() As segment
    segments = Me.segments
    With segments(UBound(segments))
        If .Type = LineSegment Then
            LastPoint = .P2
        Else
            LastPoint = .P4
        End If
    End With
End Property

' -----------------------------------------------------------------------------
' CountSubPaths (PropertyGet), source lines 220-225
' -----------------------------------------------------------------------------
Property Get CountSubPaths() As Long
    Dim iterator As LongPtr
    GdipCreatePathIter iterator, Ptr
    GdipPathIterGetSubpathCount iterator, CountSubPaths
    GdipDeletePathIter iterator
End Property

' -----------------------------------------------------------------------------
' CountNodes (PropertyGet), source lines 226-232
' -----------------------------------------------------------------------------
Property Get CountNodes() As Long
    GdipGetPointCount Ptr, CountNodes
'    Dim iterator As LongPtr
'    GdipCreatePathIter iterator, Ptr
'    GdipPathIterGetCount iterator, CountNodes
'    GdipDeletePathIter iterator
End Property

' -----------------------------------------------------------------------------
' Bounds (PropertyGet), source lines 233-236
' -----------------------------------------------------------------------------
Property Get Bounds() As RECTL
    If NeedRecalculateBounds Then CalculateBounds
    Bounds = lBounds
End Property

' -----------------------------------------------------------------------------
' CalculateBounds (Sub), source lines 237-244
' -----------------------------------------------------------------------------
Private Sub CalculateBounds()
    GdipGetPathWorldBoundsI Ptr, lBounds, 0, 0
    With lBounds
        .right = .left + .right - 1
        .bottom = .top + .bottom - 1
    End With
    NeedRecalculateBounds = False
End Sub

' -----------------------------------------------------------------------------
' pen (PropertyGet), source lines 245-247
' -----------------------------------------------------------------------------
Property Get pen() As waPen
    Set pen = lPen
End Property

' -----------------------------------------------------------------------------
' pen (PropertySet), source lines 248-250
' -----------------------------------------------------------------------------
Property Set pen(oValue As waPen)
    Set lPen = oValue
End Property

' -----------------------------------------------------------------------------
' brush (PropertyGet), source lines 252-254
' -----------------------------------------------------------------------------
Property Get brush() As waBrush
    Set brush = lBrush
End Property

' -----------------------------------------------------------------------------
' brush (PropertySet), source lines 255-257
' -----------------------------------------------------------------------------
Property Set brush(oValue As waBrush)
    Set lBrush = oValue
End Property

' -----------------------------------------------------------------------------
' Clone (Function), source lines 258-273
' -----------------------------------------------------------------------------
Public Function Clone() As waPath
    Dim cPtr As LongPtr
    GdipClonePath Ptr, cPtr
    
    Set Clone = New waPath
    Dim tmpPtr As LongPtr
    
    tmpPtr = Clone.Ptr
    Clone.Ptr = cPtr
    GdipDeletePath tmpPtr
    
    Set Clone.brush = lBrush
    Set Clone.pen = lPen
    

End Function

' -----------------------------------------------------------------------------
' AddLineSegment (Sub), source lines 274-280
' -----------------------------------------------------------------------------
Public Sub AddLineSegment(x As Long, y As Long)
    Static p As POINTL
    p.x = x
    p.y = y
    GdipAddPathLine2I Ptr, p, 1
    NeedRecalculateBounds = True
End Sub

' -----------------------------------------------------------------------------
' AddBezierSegment (Sub), source lines 281-293
' -----------------------------------------------------------------------------
Public Sub AddBezierSegment(x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long, x4 As Long, y4 As Long)
    Static p(3) As POINTL
    p(0).x = x1
    p(0).y = y1
    p(1).x = x2
    p(1).y = y2
    p(2).x = x3
    p(2).y = y3
    p(3).x = x4
    p(3).y = y4
    GdipAddPathBeziersI Ptr, p(0), 4
    NeedRecalculateBounds = True
End Sub

' -----------------------------------------------------------------------------
' ApplyMatrixTransform (Sub), source lines 294-297
' -----------------------------------------------------------------------------
Public Sub ApplyMatrixTransform(matrix As waMatrix)
    GdipTransformPath Ptr, matrix.Ptr
    NeedRecalculateBounds = True
End Sub

' -----------------------------------------------------------------------------
' StartSubPath (Sub), source lines 299-301
' -----------------------------------------------------------------------------
Public Sub StartSubPath()
    GdipStartPathFigure Ptr
End Sub

' -----------------------------------------------------------------------------
' CloseSubPath (Sub), source lines 302-304
' -----------------------------------------------------------------------------
Public Sub CloseSubPath()
    GdipClosePathFigures Ptr
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 305-308
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()

    GdipDeletePath Ptr
End Sub
