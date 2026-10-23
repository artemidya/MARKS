Private Function ReadFile(fileName As String) As String
      On Error GoTo 10
      Dim MyFile As Integer:  MyFile = FreeFile
      Open fileName For Input As #MyFile
            ReadFile = Input$(LOF(MyFile), #MyFile)
      Close #MyFile
10
End Function
