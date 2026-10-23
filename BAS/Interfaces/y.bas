Public Function y(OutPutName As String) As Long
    Dim i As Long
    i = GetIndex(OutPutName)
    y = Interfaces(i).y
End Function
