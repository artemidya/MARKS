Public Function x(OutPutName As String) As Long
    Dim i As Long
    i = GetIndex(OutPutName)
    x = Ports(i).x
End Function
