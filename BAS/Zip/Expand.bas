Public Function Expand(str As String) As String

OutString = "": pos = 0
InString = str

  'Dim FileNumber  As Integer
  'Dim FileNumber2 As Integer

    'If LenB(LoadName) = 0 Then
    '    LoadName = m_LoadFilename
    'End If
    'If LenB(SaveName) = 0 Then
    '    SaveName = m_SaveFilename
    'End If
   ' FileNumber = FreeFile
   ' Open LoadName For Binary As FileNumber
   ' FileNumber2 = FreeFile
   ' Open SaveName For Binary As FileNumber2
    LZDecompFile 'FileNumber, FileNumber2
    'Close FileNumber
    'Close FileNumber2
Expand = OutString
End Function
