Private Function isLittleEndianf() As Boolean
    'тестирует формат записи int32 littleEndian или BigEndian в памяти
    Dim numberForTest As Long, bytes(3) As Byte
    numberForTest = 38778244
    CopyMemory bytes(0), numberForTest, 4
    If bytes(0) = 132 Then isLittleEndianf = True
End Function
