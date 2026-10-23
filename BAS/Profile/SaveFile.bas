Private Function SaveFile(fileName As String, ByRef content As String) As Boolean
      On Error GoTo 10
      Dim MyFile As Integer:  MyFile = FreeFile
      Open fileName For Output As #MyFile
            Print #MyFile, content
      Close #MyFile
      SaveFile = True
10
End Function
