' =============================================================================
' Group: 02_Calculation_and_Nesting_Forms
' Source module: CorrectorForm.frm
' Procedures: 4
' =============================================================================

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 16-30
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    UserData.Load
    ListBox1.ColumnCount = 2
   
    Dim q As Long
    For q = 0 To ActiveLng.count - 1
        With ListBox1
            .AddItem
            .List(q, 0) = ActiveLng.Key(q)
            .List(q, 1) = ActiveLng.value(q)
        End With
    Next
    ListBox1.ListIndex = 0
    Label7.Caption = "ISO639=" & LCase(localization.CodeISO639)
End Sub

' -----------------------------------------------------------------------------
' CommandButton1_Click (Sub), source lines 32-92
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' ListBox1_Click (Sub), source lines 94-101
' -----------------------------------------------------------------------------
Private Sub ListBox1_Click()
    Label2.Caption = ActiveLng.Key(ListBox1.ListIndex)
    TextBox1.text = ActiveLng.value(ListBox1.ListIndex)
    TextBox1.SetFocus
    TextBox1.SelStart = 0
    TextBox1.SelLength = Len(TextBox1.text)
    If CheckBox2.value Then SendKeys "^q"
End Sub

' -----------------------------------------------------------------------------
' TextBox1_Change (Sub), source lines 102-106
' -----------------------------------------------------------------------------
Private Sub TextBox1_Change()
    ListBox1.List(ListBox1.ListIndex, 1) = TextBox1.text
    localization.ActiveLng.Item(ActiveLng.Key(ListBox1.ListIndex)) = TextBox1.text
    TextBox2.text = ""
End Sub
