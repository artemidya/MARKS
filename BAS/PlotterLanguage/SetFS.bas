Public Function SetFS(force As Long) As String
    If force = 0 Then Exit Function
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetFS = "FS" & CStr(force)
        Case DMPL
            'Exit (no force command in DMPL)
    End Select
End Function
