Private Function GlobalCrossTest(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef Curves() As curve) As Boolean
    Dim ua As Single, ub As Single
    Dim znamen As Single, chis1 As Single, chis2 As Single
    Dim x3 As Single, x4 As Single
    Dim y3 As Single, y4 As Single
    Dim q As Long, q2 As Long, q3 As Long
    
    For q = 1 To CountLineQ
        With Curves(q)
            x3 = .PositionX
            x4 = .PositionX + .With
            If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 30
            If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 30
            
            y3 = .PositionY - .Height
            y4 = .PositionY
            If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 30
            If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 30
        
            For q2 = 1 To .CountSubPath
                With .SubPath(q2)
                    x3 = .PositionX
                    x4 = .PositionX + .With
                    If x3 > x1 And x4 > x1 And x3 > x2 And x4 > x2 Then GoTo 20
                    If x3 < x1 And x4 < x1 And x3 < x2 And x4 < x2 Then GoTo 20
                    
                    y3 = .PositionY - .Height
                    y4 = .PositionY
                    If y3 > y1 And y4 > y1 And y3 > y2 And y4 > y2 Then GoTo 20
                    If y3 < y1 And y4 < y1 And y3 < y2 And y4 < y2 Then GoTo 20
                
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
                        If znamen = 0 Then GoTo 10
                    
                        ua = chis1 / znamen
                        ub = chis2 / znamen
                        If ua > 0 And ua < 1 And ub > 0 And ub < 1 Then GlobalCrossTest = True: Exit Function
10
                    Next
20
                End With
            Next
30
        End With
    Next
    GlobalCrossTest = False
End Function
