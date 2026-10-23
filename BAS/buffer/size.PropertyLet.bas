Property Let size(oValue As Long)
    count_bytes = oValue
    If oValue < 1 Then
         Err.Description = "Buffer size = 0"
         Err.Raise 400
    Else
        ReDim my_bytes(oValue - 1)
        NeedCalculateCRC = True
    End If
End Property
