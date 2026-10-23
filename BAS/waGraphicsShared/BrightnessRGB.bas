Public Function BrightnessRGB(R As Long, g As Long, b As Long) As Byte
    BrightnessRGB = CByte(0.2126 * R + 0.7152 * g + 0.0722 * b)
End Function
