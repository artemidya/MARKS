Public Function BrightnessRGBLong(RGB_ As Long) As Byte
    Dim R As Long, g As Long, b As Long
    b = RGB_ \ 65536
    g = (RGB_ - b * 65536) \ 256
    R = RGB_ - b * 65536 - g * 256
    BrightnessRGBLong = CByte(0.2126 * R + 0.7152 * g + 0.0722 * b)
End Function
