Public Function CoordinateSystemToString(cs As CoordinateSystems) As String
    Select Case cs
        Case Absolute
            CoordinateSystemToString = "ABSOLUTE"
        Case Relative
            CoordinateSystemToString = "RELATIVE"
        Case Else
            CoordinateSystemToString = "ABSOLUTE"
    End Select
End Function
