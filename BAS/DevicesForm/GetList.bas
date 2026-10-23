Private Function GetList() As Byte()
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.size = 4
    buffer1.bPut_Long Commands.GetPlotterList
    GetList = buffer1.bytes
End Function
