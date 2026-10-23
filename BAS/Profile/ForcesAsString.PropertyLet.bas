Public Property Let ForcesAsString(oValue As String)
    oValue = Trim(oValue)
    Dim arrStr() As String
    arrStr = Split(oValue, ";")
    If UBound(arrStr) > -1 Then
        Dim arr() As Long
        ReDim arr(UBound(arrStr))
        Dim q As Long
        For q = 0 To UBound(arrStr)
            arr(q) = StringToLong(arrStr(q))
        Next
    End If
    If Not CompareArrLong(arr, lForces) Then NeedSave = True
    lForces = arr
End Property
