Private Function ShotDevicePath(ByVal path As String) As String
    On Error GoTo 10
    path = UCase(path)
    path = Replace(path, "\\?\", "", , 1)
    path = Replace(path, "\\.\", "", , 1)
    path = Replace(path, "#", "\")

    path = Mid(path, 1, InStrRev(path, "\") - 1)
10
    ShotDevicePath = path
End Function
