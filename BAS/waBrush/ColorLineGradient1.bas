Property Get ColorLineGradient1() As RGBA
    Dim tmprgb(3) As Byte
    Dim colorLongPtr As LongPtr
    GdipGetLineColors lPtr, colorLongPtr
    CopyMemory tmprgb(0), colorLongPtr, 4
    
    With ColorLineGradient1
        .R = tmprgb(2)
        .g = tmprgb(1)
        .b = tmprgb(0)
        .a = tmprgb(3)
    End With
End Property
