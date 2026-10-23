Private Function GetTag(TagName As String, ByRef FromText As String) As String
      Dim s As String, e As String
      s = InStr(1, FromText, "<" & TagName & ">", vbTextCompare)
      If s <> 0 Then
            e = InStr(s + Len(TagName) + 2, FromText, "</" & TagName & ">", vbTextCompare)
            If e <> 0 Then
                  GetTag = Mid(FromText, s + Len(TagName) + 2, e - s - Len(TagName) - 2)
            End If
      End If
End Function
