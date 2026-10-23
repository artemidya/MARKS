Property Get Color() As RGBA
    Dim tmprgb(3) As Byte
    Dim colorLong As Long

    GdipGetSolidFillColor lPtr, colorLong
    CopyMemory tmprgb(0), colorLong, 4
    With Color
        .R = tmprgb(2)
        .g = tmprgb(1)
        .b = tmprgb(0)
        .a = tmprgb(3)
    End With
    
End Property
