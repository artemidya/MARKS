Private Function Pair(ByRef x As Long, ByRef y As Long, ByRef RelativeCoordinates As Boolean) As String
    If RelativeCoordinates Then
        Pair = x - Last_X & "," & y - Last_Y
    Else
        Pair = x & "," & y
    End If
    Last_X = x
    Last_Y = y
End Function
