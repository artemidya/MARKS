Private Sub BuildComSettingsString()
    Dim result As String
    result = Trim(BuildBaudPart(Combo1) & " " & _
             BuildParityPart(ComboBox2) & " " & _
             BuildDataBitsPart(ComboBox1) & " " & _
             BuildStopBitsPart(ComboBox3) & " " & _
             BuildFlowControlParts(ComboBox4))

    Dim RegData1 As New RegData
    RegData1.ComSettingsStr(PortName) = result
End Sub
