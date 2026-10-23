Private Function ToStr16bit(ByVal n As Long) As String
    Dim aa As Long, bb As Long
    aa = n \ 256
    bb = n - aa * 256
    ToStr16bit = Chr(aa) & Chr(bb)
End Function
