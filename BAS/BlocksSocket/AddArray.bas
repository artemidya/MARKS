Private Sub AddArray(length1 As Long, length2 As Long)
    Dim ms() As Byte
    ReDim ms(length1 + length2 - 1)
    
    MemMove ms(0), Remainder(0), length1
    MemMove ms(length1), LngReadBuffer(0), length2
    MemMove LngReadBuffer(0), ms(0), length1 + length2
End Sub
