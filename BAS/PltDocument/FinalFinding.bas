Private Sub FinalFinding(SortingCurves As Collection, CurrentQ As Long, PreviosPoint As POINTL)

    Dim PreviosSC As SortingCurve
    Dim CurrentSC As SortingCurve
    Dim NextSC As SortingCurve
    Dim PathIsClosed As Boolean
    
    Dim Current_FirstPoint As POINTL
    Dim Current_LastPoint As POINTL
    Dim Previos_LastPoint As POINTL
    Dim Next_FirstPoint As POINTL
    
    Set CurrentSC = SortingCurves(CurrentQ)
    Set PreviosSC = SortingCurves(CurrentQ - 1)
    With CurrentSC.path
        PathIsClosed = .closed
        Current_FirstPoint = .FirstPoint
        Current_LastPoint = .LastPoint
    End With
    Previos_LastPoint = PreviosSC.path.LastPoint
    
    Dim q As Long
    Dim a As Long, b_plus_c As Long
    Dim pp As NearestPointInfo
    Dim d1 As Long, d2 As Long
    Dim NeedReverse As Boolean
    Dim SumWay As Long
    
    Dim BEST_SumWay As Long ' BEST_SumWay = 0
    Dim BEST_Before_q As Long
    Dim BEST_Node_q As Long
    Dim BEST_NeedReverse As Boolean
    
    a = Distance(Current_FirstPoint.x, Current_FirstPoint.y, Previos_LastPoint.x, Previos_LastPoint.y)
    
    'First iteration
    Set NextSC = SortingCurves(1)
    Next_FirstPoint = NextSC.path.FirstPoint
    If PathIsClosed Then
        pp = FindNearestPointForTwoShapes(CurrentSC.path, Next_FirstPoint.x, Next_FirstPoint.y, PreviosPoint.x, PreviosPoint.y)
        b_plus_c = pp.Lenth
    Else
        d1 = Distance(Current_FirstPoint.x, Current_FirstPoint.y, PreviosPoint.x, PreviosPoint.y) + Distance(Current_LastPoint.x, Current_LastPoint.y, Next_FirstPoint.x, Next_FirstPoint.y)
        d2 = Distance(Current_LastPoint.x, Current_LastPoint.y, PreviosPoint.x, PreviosPoint.y) + Distance(Current_FirstPoint.x, Current_FirstPoint.y, Next_FirstPoint.x, Next_FirstPoint.y)
        If d1 > d2 Then
            NeedReverse = True
            b_plus_c = d2
        Else
            NeedReverse = False
            b_plus_c = d1
        End If
    End If
'   Calculate Sum Way
    SumWay = a - b_plus_c
    If SumWay > BEST_SumWay Then
        BEST_SumWay = SumWay
        BEST_Before_q = 1
        BEST_Node_q = pp.Number
        BEST_NeedReverse = NeedReverse
    End If

    For q = 2 To CurrentQ - 1
        Set NextSC = SortingCurves(q)
        Set PreviosSC = SortingCurves(q - 1)
        Next_FirstPoint = NextSC.path.FirstPoint
        Previos_LastPoint = PreviosSC.path.LastPoint
        If PathIsClosed Then
            pp = FindNearestPointForTwoShapes(CurrentSC.path, Next_FirstPoint.x, Next_FirstPoint.y, Previos_LastPoint.x, Previos_LastPoint.y)
            b_plus_c = pp.Lenth
        Else
            d1 = Distance(Current_FirstPoint.x, Current_FirstPoint.y, Previos_LastPoint.x, Previos_LastPoint.y) + Distance(Current_LastPoint.x, Current_LastPoint.y, Next_FirstPoint.x, Next_FirstPoint.y)
            d2 = Distance(Current_LastPoint.x, Current_LastPoint.y, Previos_LastPoint.x, Previos_LastPoint.y) + Distance(Current_FirstPoint.x, Current_FirstPoint.y, Next_FirstPoint.x, Next_FirstPoint.y)
            If d1 > d2 Then
                NeedReverse = True
                b_plus_c = d2
            Else
                NeedReverse = False
                b_plus_c = d1
            End If
        End If
'       Calculate Sum Way
        SumWay = a - b_plus_c
        If SumWay > BEST_SumWay Then
            BEST_SumWay = SumWay
            BEST_Before_q = q
            BEST_Node_q = pp.Number
            BEST_NeedReverse = NeedReverse
        End If
    Next

    If BEST_SumWay > 0 Then
        SortingCurves.Remove CurrentQ
        If PathIsClosed Then
            Set CurrentSC.path = CurrentSC.path.ChangeStartNode(BEST_Node_q)
        Else
            If BEST_NeedReverse Then
                CurrentSC.path.Reverse
            End If
        End If
        SortingCurves.Add CurrentSC, , BEST_Before_q
    End If
End Sub
