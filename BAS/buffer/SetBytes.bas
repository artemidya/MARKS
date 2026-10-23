Public Function SetBytes(ByRef bytes() As Byte) As Long
    my_bytes = bytes
    count_bytes = UBound(my_bytes) + 1
    NeedCalculateCRC = True
End Function
