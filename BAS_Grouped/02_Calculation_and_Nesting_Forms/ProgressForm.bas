' =============================================================================
' Group: 02_Calculation_and_Nesting_Forms
' Source module: ProgressForm.frm
' Procedures: 9
' =============================================================================

' -----------------------------------------------------------------------------
' CommandButton1_Click (Sub), source lines 18-20
' -----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    Skip = True
End Sub

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 22-30
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Skip = False
    MbClosing = False
    Break = False
    ProgressForm.PBar.Caption = ""
    ProgressForm.PBarStrip.Width = 0
    k = ProgressForm.PBar.Width
    CommandButton1.Caption = localization.ActiveLng.Item("ProgressForm.Skip") & " >>"
End Sub

' -----------------------------------------------------------------------------
' Percent (PropertyLet), source lines 31-42
' -----------------------------------------------------------------------------
Property Let Percent(p As Double)
    If p > 1 Then
        oPercent = 1
    ElseIf p < 0 Then
        oPercent = 0
    Else
        oPercent = p
    End If
    ProgressForm.PBarStrip.Width = Int(k * oPercent)
    ProgressForm.PBar.Caption = Int(oPercent * 100) & "%"
    Me.Repaint
End Property

' -----------------------------------------------------------------------------
' Percent (PropertyGet), source lines 43-45
' -----------------------------------------------------------------------------
Property Get Percent() As Double
    Percent = oPercent
End Property

' -----------------------------------------------------------------------------
' StageName (PropertyLet), source lines 46-49
' -----------------------------------------------------------------------------
Property Let StageName(txt As String)
    Label1.Caption = txt
    Me.Repaint
End Property

' -----------------------------------------------------------------------------
' EnableSkip (PropertyLet), source lines 50-55
' -----------------------------------------------------------------------------
Property Let EnableSkip(s As Boolean)
    CommandButton1.Visible = s
    If EnableSkip Then
        Skip = False
    End If
End Property

' -----------------------------------------------------------------------------
' EnableSkip (PropertyGet), source lines 57-59
' -----------------------------------------------------------------------------
Property Get EnableSkip() As Boolean
    EnableSkip = CommandButton1.Visible
End Property

' -----------------------------------------------------------------------------
' UserForm_KeyPress (Sub), source lines 61-65
' -----------------------------------------------------------------------------
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Break = True
    End If
End Sub

' -----------------------------------------------------------------------------
' UserForm_QueryClose (Sub), source lines 67-72
' -----------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not MbClosing Then
        Break = True
        Cancel = 1
    End If
End Sub
