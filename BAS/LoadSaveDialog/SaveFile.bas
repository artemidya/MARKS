Public Function SaveFile(filter As filter, Optional Directory As String, Optional fileName As String, Optional Title As String) As String
    Dim strFilter As String
    Dim lngFlags As Long
    Dim varFileName As String
    Dim hwnd As Long: hwnd = Application.AppWindow.Handle
    If IsMissing(Directory) Then Directory = ""
    If IsMissing(Title) Then Title = ""
    If IsMissing(fileName) Then fileName = ""
    Dim q As Long
    If Not Not filter.FilterRows Then
        For q = 0 To UBound(filter.FilterRows)
            strFilter = strFilter & filter.FilterRows(q).Description & vbNullChar & filter.FilterRows(q).Mask & vbNullChar
        Next
    Else
        strFilter = "(*.*)" & vbNullChar & "*.*" & vbNullChar
    End If
    varFileName = ahtCommonFileOpenSave(hwnd, lngFlags, Directory, strFilter, , , fileName, Title, False)
    If Not IsNull(varFileName) Then
        varFileName = TrimNull(varFileName)
    End If
    SaveFile = varFileName
End Function
