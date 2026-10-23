Private Function DeGetByte() As Byte

    InFilePointer& = InFilePointer& + 1
    If InPointer& = bufferSize Then
        ReDim InBuffer(bufferSize - 1)
        'Get Infile, , InBuffer
        StrToBite InBuffer
        
        BufferPointer& = 0
        DeGetByte = InBuffer(0)
        InPointer& = 1
     Else 'NOT INPOINTER&...
        DeGetByte = InBuffer(InPointer&)
        InPointer& = InPointer& + 1
    End If

End Function
