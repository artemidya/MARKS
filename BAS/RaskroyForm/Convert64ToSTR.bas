Private Function Convert64ToSTR(ByRef s As Double) As String
    Dim bytArray(7) As Byte
    CopyMemory bytArray(0), s, 8
    Convert64ToSTR = Chr(bytArray(0)) & Chr(bytArray(1)) & Chr(bytArray(2)) & Chr(bytArray(3)) & Chr(bytArray(4)) & Chr(bytArray(5)) & Chr(bytArray(6)) & Chr(bytArray(7))
End Function
