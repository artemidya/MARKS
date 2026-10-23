Public Property Get Statistic_BytesReceiving() As Long
    If LengthBIGDatagramm > 0 Then
        Statistic_BytesReceiving = CP_BIGDatagramm + CP_Datagramm
    Else
        Statistic_BytesReceiving = CP_Datagramm
    End If
End Property
