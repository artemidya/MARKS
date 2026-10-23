Private Sub DrawFigure(ByRef Figure As RCurve, ByVal x As Double, ByVal y As Double, ByVal gradus As Double, ByVal ks As Double)
    'Создаем просто массив линий с преодразованием безъе в множество прямых линий
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
    sheetCurve.CountPaths = UBound(Figure.RPaths)
    ReDim sheetCurve.Paths(1 To sheetCurve.CountPaths)
    With Figure
        For q = 1 To UBound(.RPaths)
            aa = 1
            sNX = .RPaths(q).RSegments(1).SNPx * .MaxLenght
            sNY = .RPaths(q).RSegments(1).SNPy * .MaxLenght
            ReDim sheetCurve.Paths(q).segment(1 To 1)
            sheetCurve.Paths(q).segment(1).Xpoint = sNX
            sheetCurve.Paths(q).segment(1).Ypoint = sNY
            With .RPaths(q)
                For q2 = 1 To UBound(.RSegments)
                    SCPx = .RSegments(q2).SCPx * Figure.MaxLenght
                    SCPy = .RSegments(q2).SCPy * Figure.MaxLenght
                    ECPx = .RSegments(q2).ECPx * Figure.MaxLenght
                    ECPy = .RSegments(q2).ECPy * Figure.MaxLenght
                    If q2 = UBound(.RSegments) Then
                        eNX = sheetCurve.Paths(q).segment(1).Xpoint
                        eNY = sheetCurve.Paths(q).segment(1).Ypoint
                    Else
                        eNX = .RSegments(q2 + 1).SNPx * Figure.MaxLenght
                        eNY = .RSegments(q2 + 1).SNPy * Figure.MaxLenght
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
    
    Dim LinesPaths() As LinesPath
    ReDim LinesPaths(1 To UBound(Figure.RPaths))

    maxX = -9999999999#: maxY = -9999999999#
    Dim tX As Double, tY As Double

    With sheetCurve
        For q = 1 To UBound(.Paths)
            With .Paths(q)
                For q2 = 1 To UBound(.segment)
                    With .segment(q2)
                        LinesPaths(q).CountSegment = LinesPaths(q).CountSegment + 1
                        ReDim Preserve LinesPaths(q).LineSegments(1 To LinesPaths(q).CountSegment)
                        tX = .Xpoint
                        tY = .Ypoint
                        With LinesPaths(q).LineSegments(q2)
                            .x = tX * Cos(gradus) + tY * Sin(gradus)
                            .y = tY * Cos(gradus) - tX * Sin(gradus)
                            If .x > maxX Then maxX = .x
                            If .y > maxY Then maxY = .y
                        End With
                    End With
                Next
            End With
        Next
    End With
    '=========================================================
    Dim dx As Long, dy As Long
    dx = 6
    dy = 8
    Dim x1 As Long, y1 As Long
    Dim x2 As Long, y2 As Long
    
    For q = 1 To UBound(LinesPaths)
        With LinesPaths(q)
            For q2 = 1 To .CountSegment - 1
                x1 = (.LineSegments(q2).x - maxX) * ks + x * ks
                y1 = (.LineSegments(q2).y - maxY) * ks + y * ks
                x2 = (.LineSegments(q2 + 1).x - maxX) * ks + x * ks
                y2 = (.LineSegments(q2 + 1).y - maxY) * ks + y * ks
                DrawLine LinePen, dx + x1, WindowHeigth - y1 - dy, dx + x2, WindowHeigth - y2 - dy
            Next
        End With
    Next
End Sub
