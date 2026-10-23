Private Function HasFlag(ByVal value As Long, ByVal flag As Long) As Boolean
    HasFlag = (value And flag) = flag
End Function
