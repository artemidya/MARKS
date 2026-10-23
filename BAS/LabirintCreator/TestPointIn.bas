Private Function TestPointIn(ByVal x1 As Single, ByVal y1 As Single, ByRef SubPath As SubPath) As Boolean
    TestPointIn = False
    'Проверяет принадлежность точки к фигуре
    Dim q3 As Long
    Dim x2 As Single
    Dim x3 As Single, y3 As Single
    Dim x4 As Single, y4 As Single
    Dim chis1 As Single, chis2 As Single, znamen As Single
    Dim ua As Single, ub As Single
    x2 = -999999999999999#

    With SubPath
        For q3 = 1 To SubPath.CountNodes - 1
            x3 = .PosX(q3)
            x4 = .PosX(q3 + 1)
            If x3 > x1 And x4 > x1 Then GoTo 10
            
            y3 = .PosY(q3)
            y4 = .PosY(q3 + 1)
            If y3 > y1 And y4 > y1 Then GoTo 10
            If y3 < y1 And y4 < y1 Then GoTo 10
            
            znamen = (y4 - y3) * (x2 - x1)
            chis1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
            chis2 = (x2 - x1) * (y1 - y3)
            If znamen = 0 Then GoTo 10
            
            ua = chis1 / znamen
            ub = chis2 / znamen
    
            If ua >= 0 And ua <= 1 And ub >= 0 And ub <= 1 Then
                'пересекается
                TestPointIn = Not TestPointIn
            End If
10
        Next
    End With
End Function
