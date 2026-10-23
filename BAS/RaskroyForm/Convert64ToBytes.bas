Private Sub Convert64ToBytes(ByRef bytePointer As Byte, ByVal s As Double)
    CopyMemory bytePointer, s, 8
End Sub
