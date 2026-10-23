' =============================================================================
' Group: 01_Application_and_Menu
' Source module: ThisMacroStorage.cls
' Procedures: 3
' =============================================================================

' -----------------------------------------------------------------------------
' GlobalMacroStorage_OnApplicationEvent (Sub), source lines 11-19
' -----------------------------------------------------------------------------
Private Sub GlobalMacroStorage_OnApplicationEvent(ByVal EventName As String, parameters() As Variant)
    On Error GoTo 10
    If CalculateSanP.PerimetrFormVisible Or CalculateSanP.SquareFormVisible Then
        If EventName = "SelectionChange" Then
            CalculateSanP.Break
        End If
    End If
10
End Sub

' -----------------------------------------------------------------------------
' GlobalMacroStorage_SelectionChange (Sub), source lines 20-26
' -----------------------------------------------------------------------------
Private Sub GlobalMacroStorage_SelectionChange()
    On Error GoTo 10
    If CalculateSanP.PerimetrFormVisible Or CalculateSanP.SquareFormVisible Then
        CalculateSanP.Start
    End If
10
End Sub

' -----------------------------------------------------------------------------
' GlobalMacroStorage_Start (Sub), source lines 27-36
' -----------------------------------------------------------------------------
Private Sub GlobalMacroStorage_Start()
    On Error GoTo 10
    ToolBar.CreateToolbar
    
    Dim ReData1 As New RegData
    If ReData1.AutoUpdate Then
        UpdateForm.Show 0
    End If
10
End Sub
