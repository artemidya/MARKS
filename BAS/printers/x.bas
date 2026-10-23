Public Function x(OutPutName As String) As Long
    Dim i As Long:    i = GetIndex(OutPutName)
    x = printers(i).x
End Function
