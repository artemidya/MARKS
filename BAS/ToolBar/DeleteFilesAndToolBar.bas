Private Sub DeleteFilesAndToolBar()
    On Error Resume Next
    Kill Application.GMSManager.Projects("PlotCalc").FilePath & "icons\" & "*.*" ' Kill MacID("PNG") 'for Mac
    On Error Resume Next
    RmDir Application.GMSManager.Projects("PlotCalc").FilePath & "icons"
    DeleteToolBar
End Sub
