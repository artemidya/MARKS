Public Function StringToCoordinateSystem(text As String) As CoordinateSystems
    text = UCase(text)
    Select Case text
        Case "ABSOLUTE"
            StringToCoordinateSystem = Absolute
        Case "RELATIVE"
            StringToCoordinateSystem = Relative
        Case Else
            StringToCoordinateSystem = Absolute
    End Select
End Function
