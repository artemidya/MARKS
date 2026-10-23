Private Function PtrToString(ByVal StringPtr As LongPtr) As String
    PtrToString = Space(lstrlen(StringPtr))
    lstrcpy PtrToString, StringPtr
End Function
