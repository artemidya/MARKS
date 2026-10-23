Public Property Get Statistic_BytesReceivingSize() As Long
    If LengthBIGDatagramm > 0 Then
        Statistic_BytesReceivingSize = LengthBIGDatagramm
    Else
        Statistic_BytesReceivingSize = LengthDatagramm
    End If
End Property
