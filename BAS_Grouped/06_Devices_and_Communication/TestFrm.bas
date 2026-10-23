' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: TestFrm.frm
' Procedures: 3
' =============================================================================

' -----------------------------------------------------------------------------
' USB_Events1_Change (Sub), source lines 18-20
' -----------------------------------------------------------------------------
Private Sub USB_Events1_Change()
    TextBox1.text = TextBox1.text & Timer & ": ch" & vbNewLine
End Sub

' -----------------------------------------------------------------------------
' UserForm_Activate (Sub), source lines 22-27
' -----------------------------------------------------------------------------
Private Sub UserForm_Activate()

    Set USB_Events1 = New USB_Events

    
End Sub

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 35-37
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
'    USB_Events1.Terminate
End Sub
