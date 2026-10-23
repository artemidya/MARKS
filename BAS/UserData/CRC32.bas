Private Function CRC32(ByRef bytes() As Byte) As Long
    'Вычисляет контрольную сумму байтов (файла или пакета)
    'Работает до 8Мб
    On Error Resume Next
    Dim checkSum As Long, a As Long
    For a = 0 To UBound(bytes) - 4
       checkSum = checkSum + bytes(a)
    Next
    CRC32 = checkSum
End Function
