Private Function PU(ByRef x As Long, ByRef y As Long, ByRef RelativeCoordinates As Boolean) As String
    If Last_X <> x Or Last_Y <> y Then
        PU = "U" & Pair(x, y, RelativeCoordinates)
    End If
    LastCommandIsPD = False
End Function
