Property Get closed() As Boolean
    Dim s() As segment
    s = Me.segments
    Dim n As Long
    n = UBound(s)
    With s(n)
        If .Type = CurveSegment Then
            If s(1).P1.x = .P4.x And s(1).P1.y = .P4.y Then closed = True
        Else
            If s(1).P1.x = .P2.x And s(1).P1.y = .P2.y Then closed = True
        End If
    End With
End Property
