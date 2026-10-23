Public Sub Save()
    If SaveFile(lFile, Me.content) Then NeedSave = False
End Sub
