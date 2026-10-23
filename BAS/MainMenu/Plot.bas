Public Sub Plot()
    On Error GoTo 10
'    If COM.PortIsWorking Then
'        MsgBox "The port is busy. Try again later.", vbInformation
'        Exit Sub
'    End If
    PlotterForm.Show 1
10
End Sub
