Private Sub ConvertLongToBytes(ByRef bytePointer As Byte, ByVal s As Long)
    Dim b(3) As Byte
    b(0) = s \ 16777216
    s = s - 16777216 * CLng(b(0))
    b(1) = s \ 65536
    s = s - 65536 * CLng(b(1))
    b(2) = s \ 256
    s = s - 256 * CLng(b(2))
    b(3) = s
    CopyMemory bytePointer, b(0), 4
End Sub
