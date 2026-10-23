Private Function strTo16(ByRef str) As Long
    strTo16 = CDbl(Asc(Mid(str, 1, 2))) * 256 + Asc(Mid(str, 2, 2))
End Function
