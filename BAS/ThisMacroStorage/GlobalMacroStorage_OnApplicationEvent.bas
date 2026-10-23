Private Sub GlobalMacroStorage_OnApplicationEvent(ByVal EventName As String, parameters() As Variant)
    On Error GoTo 10
    If CalculateSanP.PerimetrFormVisible Or CalculateSanP.SquareFormVisible Then
        If EventName = "SelectionChange" Then
            CalculateSanP.Break
        End If
    End If
10
End Sub
