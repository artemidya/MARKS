Private Sub EnPutByte(ByVal TheByte As Byte)

    OutFilePointer& = OutFilePointer& + 1
    If OutPointer& = bufferSize - 1 Then
        OutBuffer(bufferSize - 1) = TheByte
        BufferPointer& = 0
        'Put Outfile, , OutBuffer
         OutString = OutString & BiteToStr(OutBuffer())
        OutPointer& = 0
     Else 'NOT OUTPOINTER&...
        OutBuffer(OutPointer&) = TheByte
        OutPointer& = OutPointer& + 1
    End If

End Sub
