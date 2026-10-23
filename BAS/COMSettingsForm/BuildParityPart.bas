Private Function BuildParityPart(Combo As MSForms.ComboBox) As String
    Dim parityCode As String

    If Combo.ListIndex < 0 Then
        BuildParityPart = ""
        Exit Function
    End If

    Select Case UCase(Combo.value)
        Case "NONE": parityCode = "N"
        Case "ODD": parityCode = "O"
        Case "EVEN": parityCode = "E"
        Case "MARK": parityCode = "M"
        Case "SPACE": parityCode = "S"
        Case Else: parityCode = "N"
    End Select

    BuildParityPart = "parity=" & parityCode
End Function
