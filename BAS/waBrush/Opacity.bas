Property Get Opacity() As Byte
    Dim tmprgb(3) As Byte
    Dim Color As Long
    GdipGetSolidFillColor lPtr, Color
    CopyMemory tmprgb(0), Color, 4
    Opacity = tmprgb(3)
End Property
