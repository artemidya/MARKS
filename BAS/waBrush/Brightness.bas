Property Get Brightness() As Byte
    Dim tmprgb(3) As Byte
    Dim Color As Long
    GdipGetSolidFillColor lPtr, Color
    CopyMemory tmprgb(0), Color, 4
    Brightness = CByte(0.2126 * tmprgb(2) + 0.7152 * tmprgb(1) + 0.0722 * tmprgb(0))
End Property
