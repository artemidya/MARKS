Public Function BrightnessRGBA(rgba_ As Long) As Byte
    Dim tmprgb(3) As Byte
    CopyMemory tmprgb(0), rgba_, 4
    BrightnessRGBA = CByte(0.2126 * tmprgb(2) + 0.7152 * tmprgb(1) + 0.0722 * tmprgb(0))
End Function
