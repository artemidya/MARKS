Public Property Get CRC32() As Long
    If NeedCalculateCRC Then
        If Not Not my_bytes Then
            lCRC32 = RtlComputeCrc32(0, VarPtr(my_bytes(0)), UBound(my_bytes) + 1)
            NeedCalculateCRC = False
        End If
    End If
    CRC32 = lCRC32
End Property
