Private Sub GlobalMacroStorage_Start()
    On Error GoTo 10
    ToolBar.CreateToolbar
    
    Dim ReData1 As New RegData
    If ReData1.AutoUpdate Then
        UpdateForm.Show 0
    End If
10
End Sub
