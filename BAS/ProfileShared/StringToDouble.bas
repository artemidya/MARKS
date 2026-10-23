Public Function StringToDouble(text As String) As Double
    On Error GoTo 10
    StringToDouble = val(text)
10
End Function
