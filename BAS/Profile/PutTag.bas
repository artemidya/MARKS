Private Function PutTag(TagName As String, ByRef text As String) As String
      PutTag = "<" & TagName & ">" & text & "</" & TagName & ">" & vbNewLine
End Function
