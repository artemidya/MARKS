Private Sub CommandButton1_Click()
    TextBox2.text = "Private Sub " & localization.CodeISO639 & "()" & VBA.vbNewLine
    TextBox2.text = TextBox2.text & "   CodeISO639 = " & Chr(34) & localization.CodeISO639 & Chr(34) & VBA.vbNewLine
    TextBox2.text = TextBox2.text & "   Set ActiveLng = New DictionaryC" & VBA.vbNewLine
    TextBox2.text = TextBox2.text & "   With ActiveLng" & VBA.vbNewLine
    Dim n As Long: n = 1
    Dim q As Long
    Dim SubLines As String
    
    If CheckBox1.value Then
        Dim chrW2 As String, q2 As Long, b As String
        Dim chrW2Line As String
        For q = 0 To ListBox1.ListCount - 1
            chrW2 = ""
            chrW2Line = ""
            For q2 = 1 To Len(localization.ActiveLng.value(q))
                b = Mid(localization.ActiveLng.value(q), q2, 1)
                chrW2Line = chrW2Line & "ChrW(" & AscW(b) & ")"
                If q2 < Len(localization.ActiveLng.value(q)) Then
                    If Len(chrW2Line) > 900 Then
                        chrW2Line = chrW2Line & " & _" & VBA.vbNewLine
                        chrW2 = chrW2 & chrW2Line
                        chrW2Line = ""
                    Else
                        chrW2Line = chrW2Line & " & "
                    End If
                End If
            Next
            chrW2 = chrW2 & chrW2Line
            SubLines = SubLines & "       .Add " & Chr(34) & localization.ActiveLng.Key(q) & Chr(34) & ", " & chrW2 & VBA.vbNewLine
            
            If Len(SubLines) > 30000 Then
                TextBox2.text = TextBox2.text & SubLines
                TextBox2.text = TextBox2.text & "   End With" & VBA.vbNewLine
                n = n + 1
                TextBox2.text = TextBox2.text & "   " & localization.CodeISO639 & n & VBA.vbNewLine
                TextBox2.text = TextBox2.text & "End Sub" & VBA.vbNewLine
                
                TextBox2.text = TextBox2.text & "Private Sub " & localization.CodeISO639 & n & "()" & VBA.vbNewLine
                TextBox2.text = TextBox2.text & "   With ActiveLng" & VBA.vbNewLine
                SubLines = ""
            End If
            
        Next
        TextBox2.text = TextBox2.text & SubLines
    Else
        For q = 0 To ListBox1.ListCount - 1
            TextBox2.text = TextBox2.text & "       .Add " & Chr(34) & localization.ActiveLng.Key(q) & Chr(34) & ", " & Chr(34) & localization.ActiveLng.value(q) & Chr(34) & VBA.vbNewLine
        Next
    End If
    
    ' .Add "SettingsForm.Caption", ChrW(20276) & ChrW(20387) & ChrW(30340) & ChrW(80) & ChrW(108) & ChrW(111) & ChrW(116) & ChrW(67) & ChrW(97) & ChrW(108) & ChrW(99)
    
    TextBox2.text = TextBox2.text & "   End With" & VBA.vbNewLine
    TextBox2.text = TextBox2.text & "End Sub"
    
    Dim d As New MSForms.DataObject
    d.SetText TextBox2.text
    d.PutInClipboard
    Label6.Caption = "ClipBoard Size=" & Len(TextBox2.text)
End Sub
