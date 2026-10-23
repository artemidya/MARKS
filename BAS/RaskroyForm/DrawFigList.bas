Private Sub DrawFigList(ByRef ks As Double)
    'NumberOfMaxFigure
    On Error Resume Next
    Dim q As Long, q2 As Long
    Dim sNX As Double, sNY As Double
    Dim eNX As Double, eNY As Double
    Dim SCPx As Double, SCPy As Double
    Dim ECPx As Double, ECPy As Double
    Dim minY As Double, maxY As Double
    Dim minX As Double, maxX As Double
    Dim aa As Long
    Dim sheetCurve As PLCurve
    Dim maxRecurse As Long
    Dim a As Double, a1 As Double, a2 As Double
    Dim da1 As Double, da2 As Double
    minX = 999999999999#
    minY = 999999999999#
    sheetCurve.CountPaths = UBound(RCurves(NumberOfMaxFigure).RPaths)
    ReDim sheetCurve.Paths(1 To sheetCurve.CountPaths)
    With RCurves(NumberOfMaxFigure)
        For q = 1 To UBound(.RPaths)
            aa = 1
            sNX = .RPaths(q).RSegments(1).SNPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
            sNY = .RPaths(q).RSegments(1).SNPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
            ReDim sheetCurve.Paths(q).segment(1 To 1)
            sheetCurve.Paths(q).segment(1).Xpoint = sNX
            sheetCurve.Paths(q).segment(1).Ypoint = sNY
            With .RPaths(q)
                For q2 = 1 To UBound(.RSegments)
                    SCPx = .RSegments(q2).SCPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    SCPy = .RSegments(q2).SCPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    ECPx = .RSegments(q2).ECPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    ECPy = .RSegments(q2).ECPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    If q2 = UBound(.RSegments) Then
                        eNX = sheetCurve.Paths(q).segment(1).Xpoint
                        eNY = sheetCurve.Paths(q).segment(1).Ypoint
                    Else
                        eNX = .RSegments(q2 + 1).SNPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                        eNY = .RSegments(q2 + 1).SNPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    End If
                    a1 = Angel(sNX, sNY, SCPx, SCPy)
                    a = Angel(SCPx, SCPy, ECPx, ECPy)
                    a2 = Angel(ECPx, ECPy, eNX, eNY)
                    da1 = Abs(DeltaAngle(a, a1))
                    da2 = Abs(DeltaAngle(a2, a))
                    If SCPx = sNX And SCPy = sNY Then da1 = 0
                    If ECPx = eNX And ECPy = eNY Then da2 = 0
                    a = da1 + da2
                    If a < 0.01 Then
                        maxRecurse = 0
                    ElseIf a < 0.05 Then
                        maxRecurse = 1
                    ElseIf a < 0.26 Then
                        maxRecurse = 2
                    ElseIf a < 0.9 Then
                        maxRecurse = 3
                    ElseIf a < 2.2 Then
                        maxRecurse = 4
                    Else
                        maxRecurse = 5
                    End If
                    recursive_bezier sNX, sNY, SCPx, SCPy, ECPx, ECPy, eNX, eNY, 1, sheetCurve.Paths(q), aa, maxRecurse, maxX, minX, maxY, minY
                    sNX = eNX
                    sNY = eNY
                Next
            End With
        Next
    End With
    'Теперь нарисовать лист
    Dim dx As Long, dy As Long
    dx = 6
    dy = 8
    For q = 1 To UBound(sheetCurve.Paths)
        With sheetCurve.Paths(q)
            For q2 = 1 To UBound(.segment)
                .segment(q2).Ypoint = -dy + WindowHeigth - .segment(q2).Ypoint + minY
                .segment(q2).Xpoint = dx + .segment(q2).Xpoint - minX
            Next
        End With
    Next
    Dim lines() As LinesString
    ReDim lines(1 To WindowHeigth) As LinesString
    Dim y As Long
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    
    For y = WindowHeigth - (maxY - minY) - dy To WindowHeigth
        For q = 1 To UBound(sheetCurve.Paths)
            With sheetCurve.Paths(q)
                For q2 = 1 To UBound(.segment) - 1
                    y1 = .segment(q2).Ypoint
                    y2 = .segment(q2 + 1).Ypoint
                    If y > y1 And y > y2 Then GoTo nx
                    If y < y1 And y < y2 Then GoTo nx
                    If y2 = y1 Then GoTo nx
                    x1 = .segment(q2).Xpoint
                    x2 = .segment(q2 + 1).Xpoint
                    lines(y).CountX = lines(y).CountX + 1
                    ReDim Preserve lines(y).x(1 To lines(y).CountX)
                    lines(y).x(lines(y).CountX) = (x2 - x1) * (y - y1) / (y2 - y1) + x1
nx:
                Next
            End With
        Next
    Next
    For q = 1 To UBound(lines)
        With lines(q)
            SortData .x()
        End With
    Next
    For q = 1 To UBound(lines)
        With lines(q)
            For q2 = 1 To .CountX Step 2
                DrawLine LinesPaperPen, .x(q2), q, .x(q2 + 1), q
                DrawLine LinesShadowPen, .x(q2) + 2, q + 2, .x(q2 + 1) + 2, q + 2
            Next
        End With
    Next
End Sub
