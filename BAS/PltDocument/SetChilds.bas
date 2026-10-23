Private Sub SetChilds(CurvesForSetChilds As Collection)

    Dim q As Long, q2 As Long
    Dim MySortingCurve As SortingCurve, MySortingCurve2 As SortingCurve
    Dim left As Long, right As Long, top As Long, bottom As Long
    Dim NeedExit As Boolean

    q = 1
    Do While q < CurvesForSetChilds.count
        Set MySortingCurve = CurvesForSetChilds(q)
        
        With MySortingCurve.path.Bounds
            left = .left
            right = .right
            top = .top
            bottom = .bottom
        End With
        
        NeedExit = False
        
        For q2 = q + 1 To CurvesForSetChilds.count
            Set MySortingCurve2 = CurvesForSetChilds(q2)
            
            With MySortingCurve2.path.Bounds
                If left >= .left Then
                    If right <= .right Then
                        If top >= .top Then
                            If bottom <= .bottom Then
                                MySortingCurve2.Childs.Add MySortingCurve
                                CurvesForSetChilds.Remove q
                                q = q - 1
                                NeedExit = True
                            End If
                        End If
                    End If
                End If
            End With
            
            If NeedExit Then Exit For
        Next
        q = q + 1
    Loop

End Sub
