Private Sub FlushBuffer()

  Dim WriteBuffer() As Byte

    If OutPointer& > 0 Then
        ReDim WriteBuffer(OutPointer - 1)
        CopyMemory WriteBuffer(0), OutBuffer(0), OutPointer
        'Put Outfile, , WriteBuffer
        OutString = OutString & BiteToStr(WriteBuffer())
     Else 'NOT OUTPOINTER&...
    End If

End Sub
