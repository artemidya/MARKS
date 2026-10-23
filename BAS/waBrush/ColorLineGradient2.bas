Property Get ColorLineGradient2() As RGBA
    Dim tmprgb(7) As Byte
    Dim colorLongPtr As LongPtr
    GdipGetLineColors lPtr, colorLongPtr
    CopyMemory tmprgb(0), colorLongPtr, 8

    With ColorLineGradient2
        .R = tmprgb(6)
        .g = tmprgb(5)
        .b = tmprgb(4)
        .a = tmprgb(7)
    End With
End Property
