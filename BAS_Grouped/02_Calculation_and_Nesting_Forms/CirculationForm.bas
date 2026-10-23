' =============================================================================
' Group: 02_Calculation_and_Nesting_Forms
' Source module: CirculationForm.frm
' Procedures: 7
' =============================================================================

' -----------------------------------------------------------------------------
' Command4_Click (Sub), source lines 11-19
' -----------------------------------------------------------------------------
Private Sub Command4_Click()
    On Error GoTo 10
    
    MakeDuplicate ConvertLng(Text1.text)
    
    Unload Me
    RaskroyForm.Show 1
10
End Sub

' -----------------------------------------------------------------------------
' Command5_Click (Sub), source lines 21-23
' -----------------------------------------------------------------------------
Private Sub Command5_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' Text1_Change (Sub), source lines 25-33
' -----------------------------------------------------------------------------
Private Sub Text1_Change()
    If ActiveSelection.shapes.count > 0 Then
        If ConvertLng(Text1.text) > 0 And ConvertLng(Text1.text) < (1000 \ ActiveSelection.shapes.count + 1) Then
            Command4.Enabled = True
        Else
            Command4.Enabled = False
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 35-55
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    UserData.Load
    If UserData.NumberOfCopies < 1 Then
        Text1.text = 10
    Else
        Text1.text = UserData.NumberOfCopies
    End If
    Text1.SelStart = 0
    Text1.SelLength = 10
    
    If ActiveSelection.shapes.count < 1 Then
        MsgBox localization.ActiveLng.Item("CirculationForm.MsgBox"), vbOKOnly Or vbInformation, "PlotCalc"
        Unload Me
    End If
    
    Command4.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton1.Caption")
    Command5.Caption = localization.ActiveLng.Item("SettingsForm.Command5.Caption")
    
    Label1.Caption = localization.ActiveLng.Item("CirculationForm.Label1.Caption")
    
End Sub

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 57-59
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    UserData.NumberOfCopies = ConvertLng(Text1.text)
End Sub

' -----------------------------------------------------------------------------
' MakeDuplicate (Sub), source lines 60-83
' -----------------------------------------------------------------------------
Private Sub MakeDuplicate(n As Long)
    If n > 1 Then
        Application.Optimization = True
        Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("CommandBars.ToolTipText1_5") & " (" & n & ")"
        Dim q As Long
        Dim OrigSelection As ShapeRange
        Set OrigSelection = ActiveSelectionRange
        Dim dup() As ShapeRange
        ReDim dup(n - 2)
        
        For q = 0 To n - 2
            Set dup(q) = OrigSelection.Duplicate
            dup(q).Move 0, OrigSelection.SizeHeight * (q + 1)
        Next
        
        OrigSelection.AddToSelection
        For q = 0 To n - 2
            dup(q).AddToSelection
        Next
        Application.ActiveDocument.EndCommandGroup
        Application.Optimization = False
        Application.Refresh
    End If
End Sub

' -----------------------------------------------------------------------------
' ConvertLng (Function), source lines 84-90
' -----------------------------------------------------------------------------
Private Function ConvertLng(t As String) As Long
    On Error GoTo 10
    ConvertLng = CLng(t)
    Exit Function
10
    Command4.Enabled = False
End Function
