Private Function FindBestAngel(path As waPath, LastAngel As Double) As Boolean
    Dim q As Long
    Dim a As Double, da As Double, BEST_a As Double, BEST_q As Long
    BEST_a = 99999 'Max Val
    Dim segments() As segment
    segments = path.segments

    For q = 1 To UBound(segments)
        With segments(q)
            a = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
            da = DeltaAngleAbs(LastAngel, a)
            If da = 0 Then
                BEST_q = q
                q = UBound(segments)
                GoTo 10
            End If
            If da < BEST_a Then
                BEST_a = da
                BEST_q = q
            End If
10
        End With
    Next
    
    'Change start point
    If BEST_q > 1 Then
        Set path = path.ChangeStartNode(BEST_q)
        FindBestAngel = True
    End If
End Function
