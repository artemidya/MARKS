Private Sub AddLine(ByVal x As Single, ByVal y As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef Curves() As curve)
    'Процедура добавляет линию!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    CountLineQ = CountLineQ + 1
    ReDim Preserve Curves(CountLineQ)
    With Curves(CountLineQ)
        .CountSubPath = 1
        ReDim .SubPath(1)
        If x < x2 Then
            .PositionX = x
            .With = x2 - x
        Else
            .PositionX = x2
            .With = x - x2
        End If
        .SubPath(1).PositionX = .PositionX
        .SubPath(1).With = .With
        If y > y2 Then
            .PositionY = y
            .Height = y - y2
        Else
            .PositionY = y2
            .Height = y2 - y
        End If
        .SubPath(1).PositionY = .PositionY
        .SubPath(1).Height = .Height
        With .SubPath(1)
            .CountNodes = 2
            ReDim .PosX(2)
            ReDim .PosY(2)
            .PosX(1) = x
            .PosY(1) = y
            .PosX(2) = x2
            .PosY(2) = y2
        End With
    End With
End Sub
