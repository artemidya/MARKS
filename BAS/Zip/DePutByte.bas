Private Sub DePutByte(ByVal TheByte As Byte)

  Dim WriteBuffer() As Byte

    OutFilePointer& = OutFilePointer& + 1
    If OutPointer& = bufferSize - 1 Then
        OutBuffer(bufferSize - 1) = TheByte
        ReDim WriteBuffer(OutPointer)
        CopyMemory WriteBuffer(0), OutBuffer(0), OutPointer + 1
        'Put Outfile, , WriteBuffer
        OutString = OutString & BiteToStr(WriteBuffer())
        OutPointer& = 0
     Else 'NOT OUTPOINTER&...
        OutBuffer(OutPointer&) = TheByte
        OutPointer& = OutPointer& + 1
    End If

End Sub
