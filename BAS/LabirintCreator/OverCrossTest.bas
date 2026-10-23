Private Function OverCrossTest(ByRef xp As Single, ByRef yp As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef SecError As Boolean, ByRef Curves() As curve) As Boolean
    SecError = False
    'Процедура для нахождения пересечений при окончании работы.
    'xP возвращается X пересечения
    'yP возвращается Y пересечения
    xp = 0
    yp = 0
    OverCrossTest = False
    Dim ua As Single, ub As Single
    Dim znamen As Single, chis1 As Single, chis2 As Single
    Dim x3 As Single, x4 As Single
    Dim y3 As Single, y4 As Single
    Dim minR As Single, R As Single
    Dim r1 As Single, R2 As Single
    Dim xPtemp As Single, yPtemp As Single
    minR = 99999999999999#
    Dim q As Long, q2 As Long, q3 As Long
    Dim qSec As Long, q2Sec As Long
    
    For q = 1 To CountLineQ
        x3 = Curves(q).PositionX
        x4 = Curves(q).PositionX + Curves(q).With
        If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 30
        If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 30
        
        y3 = Curves(q).PositionY - Curves(q).Height
        y4 = Curves(q).PositionY
        If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 30
        If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 30
        
        For q2 = 1 To Curves(q).CountSubPath
            x3 = Curves(q).SubPath(q2).PositionX
            x4 = Curves(q).SubPath(q2).PositionX + Curves(q).SubPath(q2).With
            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 20
            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 20
            
            y3 = Curves(q).SubPath(q2).PositionY - Curves(q).SubPath(q2).Height
            y4 = Curves(q).SubPath(q2).PositionY
            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 20
            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 20
        
            For q3 = 1 To Curves(q).SubPath(q2).CountNodes - 1
                x3 = Curves(q).SubPath(q2).PosX(q3)
                x4 = Curves(q).SubPath(q2).PosX(q3 + 1)
                If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 10
                If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 10
                
                y3 = Curves(q).SubPath(q2).PosY(q3)
                y4 = Curves(q).SubPath(q2).PosY(q3 + 1)
                If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 10
                If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 10
                
                znamen = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
                chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
                chis2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
                If znamen = 0 Then GoTo 10
                
                ua = chis1 / znamen
                ub = chis2 / znamen
                If ua >= 0 And ua <= 1 And ub >= 0 And ub <= 1 Then
                    xPtemp = x1 + ua * (x2 - x1)
                    yPtemp = y1 + ua * (y2 - y1)
                    'R = Sqr((xPtemp - X1) * (xPtemp - X1) + (yPtemp - Y1) * (yPtemp - Y1))
                    R = (xPtemp - x1) * (xPtemp - x1) + (yPtemp - y1) * (yPtemp - y1)
                    If R < minR And R <> 0 Then
                        If q > CountCurves Then
                            xp = xPtemp
                            yp = yPtemp
                        Else
                            r1 = (xPtemp - x3) * (xPtemp - x3) + (yPtemp - y3) * (yPtemp - y3)
                            R2 = (xPtemp - x4) * (xPtemp - x4) + (yPtemp - y4) * (yPtemp - y4)
                            If r1 < R2 Then
                                xp = x3
                                yp = y3
                            Else
                                xp = x4
                                yp = y4
                            End If
                        End If
                        minR = R
                        OverCrossTest = True
                        qSec = q
                        q2Sec = q2
                    End If
                End If
10
            Next
20
        Next
30
    Next
    If OverCrossTest = True Then
        If TestPointIn(x1, y1, Curves(qSec).SubPath(q2Sec)) = True Then
            SecError = True
        End If
        If GlobalCrossTest(x1, y1, xp, yp, Curves) = True Then
            SecError = True
        End If
    End If
End Function
