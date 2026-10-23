Private Function EnGetByte() As Byte

    InFilePointer& = InFilePointer& + 1
    If InPointer& = bufferSize Then
        'Get Infile, , InBuffer
        StrToBite InBuffer
        EnGetByte = InBuffer(0)
        InPointer& = 1
     Else 'NOT INPOINTER&...
        EnGetByte = InBuffer(InPointer&)
        InPointer& = InPointer& + 1
    End If

End Function
