Private Function CrossPoints(ByRef XS() As Double, ByRef YS() As Double, ByRef Curves() As curve, ByRef CountCurves As Long) As Long
    Dim q1 As Long, q2 As Long, q3 As Long
    Dim a1 As Long, a2 As Long, a3 As Long
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    Dim x3 As Double, x4 As Double
    Dim y3 As Double, y4 As Double
    
    Dim ua As Double, ub As Double
    Dim znamen As Double, chis1 As Double, chis2 As Double
    
    For q1 = 1 To CountCurves - 1
        With Curves(q1)
            For q2 = 1 To .CountSubPath
                For q3 = 1 To .SubPath(q2).CountNodes - 1
                    With .SubPath(q2)
                        x1 = .PosX(q3)
                        x2 = .PosX(q3 + 1)
                        y1 = .PosY(q3)
                        y2 = .PosY(q3 + 1)
                    End With
                        '================================
        
                    For a1 = q1 + 1 To CountCurves
                        With Curves(a1)
                            x3 = .PositionX
                            x4 = .PositionX + .With
                            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 30
                            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 30
        
                            y3 = .PositionY - .Height
                            y4 = .PositionY
                            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 30
                            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 30
                            For a2 = 1 To .CountSubPath
                                With .SubPath(a2)
                                    x3 = .PositionX
                                    x4 = .PositionX + .With
                                    If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 20
                                    If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 20
                                    y3 = .PositionY - .Height
                                    y4 = .PositionY
                                    If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 20
                                    If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 20
                                    For a3 = 1 To .CountNodes - 1
                                        x3 = .PosX(a3)
                                        x4 = .PosX(a3 + 1)
                                        If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 10
                                        If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 10
                                        y3 = .PosY(a3)
                                        y4 = .PosY(a3 + 1)
                                        If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 10
                                        If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 10
                                        '---Проверка на пересечение---
                                        znamen = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
                                        If znamen = 0 Then GoTo 10
                                        chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
                                        chis2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
                                        ua = chis1 / znamen
                                        ub = chis2 / znamen
                                        If ua > 0 And ua <= 1 And ub >= 0 And ub < 1 Then
                                            CrossPoints = CrossPoints + 1
                                            ReDim Preserve XS(CrossPoints)
                                            ReDim Preserve YS(CrossPoints)
                                            XS(CrossPoints) = x1 + ua * (x2 - x1)
                                            YS(CrossPoints) = y1 + ua * (y2 - y1)
                                        End If
10
                                    Next
20
                                End With
                            Next
30
                        End With
                    Next
                Next
            Next
        End With
    Next
End Function
