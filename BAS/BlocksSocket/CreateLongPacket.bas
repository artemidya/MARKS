Public Function CreateLongPacket(ByRef ByteBuffer() As Byte, ByVal lengthByteBuffer As Long, Optional ByVal sp As Long = 0) As Long()    'Do 2 Gb
    Dim lengthLongBuffer As Long
    lengthLongBuffer = (lengthByteBuffer + 3) \ 4 + 3
    Dim longBuffer() As Long
    ReDim longBuffer(lengthLongBuffer - 1)
    MemMove longBuffer(3), ByteBuffer(sp), lengthByteBuffer
    longBuffer(1) = htonl(lengthByteBuffer)
    Dim ZamTerminator As Long
    'ZamTerminator = 0
    Do While FindLong(ZamTerminator, longBuffer)
        ZamTerminator = NewZamTerminator
    Loop
    If ZamTerminator = 0 Then
        'No replace
        ZamTerminator = NewZamTerminator
        SndFlags.NeedReplaceFlag = 0
    Else
        'Replace
        Replace_0_ToZamTerminatorLong ZamTerminator, longBuffer
        SndFlags.NeedReplaceFlag = 1
    End If
    MemMove longBuffer(1), FlagsToByte(SndFlags), 1
    longBuffer(2) = ZamTerminator
    CreateLongPacket = longBuffer
End Function
