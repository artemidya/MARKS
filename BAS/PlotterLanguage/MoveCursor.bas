Public Function MoveCursor(x As Long, y As Long, Optional velocity As Long = 50) As String
    velocity = velocity \ 10
    If velocity < 1 Then velocity = 1
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            MoveCursor = "PA;VS" & velocity & ";PU" & CStr(x) & "," & CStr(y) & ";"
        Case DMPL
            MoveCursor = "AV" & velocity & ",U" & CStr(x) & "," & CStr(y) & ","
    End Select
End Function
