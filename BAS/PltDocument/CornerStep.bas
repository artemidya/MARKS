Private Sub CornerStep(LastAngel As Double)
    Dim q As Long
    Dim path As waPath
    Dim a As Double, aRev As Double
    
    With Curves
        For q = 1 To .count
            Set path = Curves(q)
            With path
                If .closed Then
                    a = AngelStart(path)
                    aRev = AngelEndRev(path)
                    If DeltaAngleAbs(LastAngel, a) > DeltaAngleAbs(LastAngel, aRev) Then .Reverse
                    'If little
                    If .Bounds.bottom - .Bounds.top <= 30000 Then
                        If FindBestAngel(path, LastAngel) Then
                            Curves.Add path, , , q
                            Curves.Remove q
                        End If
                    End If
                End If
            End With
            LastAngel = AngelEnd(path)
        Next
    End With
End Sub
