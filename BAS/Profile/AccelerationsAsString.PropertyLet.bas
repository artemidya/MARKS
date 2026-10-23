Public Property Let AccelerationsAsString(oValue As String)
    oValue = Trim(oValue)
    Dim arrStr() As String
    arrStr = Split(oValue, ";")
    If UBound(arrStr) > -1 Then
        Dim arr() As Long
        ReDim arr(UBound(arrStr))
        Dim q As Long
        For q = 0 To UBound(arrStr)
            arr(q) = StringToLong(StringToDouble(arrStr(q)) * 10)
        Next
    End If
    If Not CompareArrLong(arr, lAccelerations) Then NeedSave = True
    lAccelerations = arr
End Property
