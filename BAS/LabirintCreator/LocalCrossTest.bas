Private Function LocalCrossTest(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef SubPath As SubPath) As Boolean
    Dim ua As Single, ub As Single
    Dim znamen As Single, chis1 As Single, chis2 As Single
    Dim x3 As Single, x4 As Single
    Dim y3 As Single, y4 As Single
    Dim q3 As Long
    With SubPath
        For q3 = 1 To .CountNodes - 1
            x3 = .PosX(q3)
            x4 = .PosX(q3 + 1)
            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 10
            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 10
            
            y3 = .PosY(q3)
            y4 = .PosY(q3 + 1)
            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 10
            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 10
            
            znamen = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
            chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
            chis2 = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
            
            If znamen = 0 Then
                If chis1 = 0 Or chis2 = 0 Then LocalCrossTest = True: GoTo 20
            Else
                ua = chis1 / znamen
                ub = chis2 / znamen
                If ua > 0 And ua < 1 And ub > 0 And ub < 1 Then LocalCrossTest = True: GoTo 20
            End If
10
        Next
        LocalCrossTest = False
20
    End With
End Function
