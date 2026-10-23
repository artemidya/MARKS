Private Sub EnableCommand1()
    If Connected Then
        If Not Not RCurves Then
            If NotNest Then
                Command1.Enabled = False
            Else
                Command1.Enabled = True
            End If
        End If
    End If
End Sub
