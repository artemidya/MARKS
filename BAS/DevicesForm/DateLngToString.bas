Private Function DateLngToString(DateLng As Long) As String
    On Error GoTo 10
    Dim t As String
    t = CStr(DateLng)
    DateLngToString = right(t, 2) & "." & Mid(t, 5, 2) & "." & left(t, 4)
10
End Function
