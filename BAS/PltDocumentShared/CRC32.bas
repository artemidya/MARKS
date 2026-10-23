Private Function CRC32(ByRef txtForSignature As String) As String
    'Вычисляет контрольную сумму пакета
    'Работает до 8Мб
    On Error Resume Next
    Dim checkSum As Long, a As Long
    For a = 1 To Len(txtForSignature)
       checkSum = checkSum + Asc(Mid(txtForSignature, a, 1))
    Next
    CRC32 = checkSum
End Function
