Public Sub bPut_StringASCII(ByVal value As String)
    If cp + Len(value) > count_bytes Then
        count_bytes = cp + Len(value)
        ReDim Preserve my_bytes(count_bytes - 1)
        bPut_StringASCII value
    Else
        Dim q As Long
        For q = 1 To Len(value)
            my_bytes(cp + q - 1) = Asc(Mid(value, q, 1))
        Next
        cp = cp + Len(value)
        NeedCalculateCRC = True
    End If
End Sub
