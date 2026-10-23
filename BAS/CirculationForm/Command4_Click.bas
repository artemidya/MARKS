Private Sub Command4_Click()
    On Error GoTo 10
    
    MakeDuplicate ConvertLng(Text1.text)
    
    Unload Me
    RaskroyForm.Show 1
10
End Sub
