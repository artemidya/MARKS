Private Sub OpenBuffer()

    ReDim InBuffer(bufferSize - 1) As Byte
    ReDim OutBuffer(bufferSize - 1) As Byte
    InLen& = DataLen&
   ' Get Infile, , InBuffer
    StrToBite InBuffer
    InPointer& = 0
    InFilePointer& = 0
    OutPointer& = 0
    OutFilePointer& = 0

End Sub
