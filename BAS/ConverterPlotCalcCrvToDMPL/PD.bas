Private Function PD(ByRef x As Long, ByRef y As Long, ByRef RelativeCoordinates As Boolean) As String
    If Last_X <> x Or Last_Y <> y Then
        If LastCommandIsPD Then
            PD = Pair(x, y, RelativeCoordinates)
        Else
            PD = "D" & Pair(x, y, RelativeCoordinates)
        End If
        LastCommandIsPD = True
    End If
End Function
