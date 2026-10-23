Private Function CalculatePlotterBoundingBox(Curves As Collection) As RECTL
    Dim q As Long
    Dim p As waPath
    
    Dim top As Long: top = 2147483647
    Dim left As Long: left = 2147483647
    Dim right As Long: right = -2147483647
    Dim bottom As Long: bottom = -2147483647
    
    For q = 1 To Curves.count
        Set p = Curves(q)
        With p.Bounds
            If .top < top Then top = .top
            If .bottom > bottom Then bottom = .bottom
            If .right > right Then right = .right
            If .left < left Then left = .left
        End With
    Next
    With CalculatePlotterBoundingBox
        .top = bottom
        .bottom = top
        .left = left
        .right = right
    End With
    
End Function
