' =============================================================================
' Group: 09_Licensing_and_Updates
' Source module: UpdateForm.frm
' Procedures: 2
' =============================================================================

' -----------------------------------------------------------------------------
' UpdateCenter1_DownloadPercent (Sub), source lines 15-17
' -----------------------------------------------------------------------------
Private Sub UpdateCenter1_DownloadPercent(Percent As Long)
    Label1.Caption = Percent & "%"
End Sub

' -----------------------------------------------------------------------------
' UserForm_Activate (Sub), source lines 19-27
' -----------------------------------------------------------------------------
Private Sub UserForm_Activate()
    On Error GoTo 10
    Set UpdateCenter1 = New UpdateCenter
    If UpdateCenter1.version > UserData.PlotCalcVersion Then
        UpdateCenter1.UpdatePlotCalc
    End If
    Unload Me
10
End Sub
