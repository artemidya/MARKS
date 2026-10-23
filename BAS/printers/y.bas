Public Function y(OutPutName As String) As Long
    Dim i As Long:    i = GetIndex(OutPutName)
    y = printers(i).y
End Function
