' =============================================================================
' Group: 09_Licensing_and_Updates
' Source module: ActivationForm.frm
' Procedures: 13
' =============================================================================

' -----------------------------------------------------------------------------
' Command1_MouseMove (Sub), source lines 15-17
' -----------------------------------------------------------------------------
Private Sub Command1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label2.Font.Underline = True
End Sub

' -----------------------------------------------------------------------------
' Text1_KeyDown (Sub), source lines 18-20
' -----------------------------------------------------------------------------
Private Sub Text1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' Text1_MouseMove (Sub), source lines 21-23
' -----------------------------------------------------------------------------
Private Sub Text1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label2.Font.Underline = True
End Sub

' -----------------------------------------------------------------------------
' UserForm_Activate (Sub), source lines 25-28
' -----------------------------------------------------------------------------
Private Sub UserForm_Activate()
    Label1.Caption = localization.ActiveLng.Item("ActivationForm.Label1.Caption1") & " " & HoursAndMinets(ParentForm.Time_for_work \ 60 + 1) & localization.ActiveLng.Item("ActivationForm.Label1.Caption2") & vbNewLine & vbNewLine _
                    & localization.ActiveLng.Item("ActivationForm.Label1.Caption3")
End Sub

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 30-41
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    IsEditingMyFields = True

    Caption = localization.ActiveLng.Item("ActivationForm.Caption")
    Command1.Caption = localization.ActiveLng.Item("ActivationForm.Command1.Caption")
    Label2.Caption = localization.ActiveLng.Item("ActivationForm.Label2.Caption")
    Label2.left = 158 - Label2.Width \ 2
    Label2.ControlTipText = localization.ActiveLng.Item("ActivationForm.BuyURL")
    
    Text1.text = UCase(Trim(UserData.CodeOfActivation))
    IsEditingMyFields = False
End Sub

' -----------------------------------------------------------------------------
' Command1_Click (Sub), source lines 42-51
' -----------------------------------------------------------------------------
Private Sub Command1_Click()
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.size = 44
    buffer1.bPut_Long Commands.ActivationToActCode
    Dim code As String * 20
    code = Text1.text
    buffer1.bPut_String code
    ParentForm.BlocksSocket1.Send buffer1.bytes
End Sub

' -----------------------------------------------------------------------------
' Label2_Click (Sub), source lines 52-56
' -----------------------------------------------------------------------------
Private Sub Label2_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", localization.ActiveLng.Item("ActivationForm.BuyURL"), "", "", 4
10:
End Sub

' -----------------------------------------------------------------------------
' Label1_MouseMove (Sub), source lines 57-59
' -----------------------------------------------------------------------------
Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label2.Font.Underline = True
End Sub

' -----------------------------------------------------------------------------
' Label2_MouseMove (Sub), source lines 60-66
' -----------------------------------------------------------------------------
Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Label2.Font.Underline Then
        Label2.Font.Underline = False
        DoEvents
    End If
     Cursor.SetCursor2 Cursor.IDC_HAND
End Sub

' -----------------------------------------------------------------------------
' UserForm_MouseMove (Sub), source lines 67-69
' -----------------------------------------------------------------------------
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label2.Font.Underline = True
End Sub

' -----------------------------------------------------------------------------
' Text1_Change (Sub), source lines 70-78
' -----------------------------------------------------------------------------
Private Sub Text1_Change()
    If Len(Text1.text) = 20 Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
    If IsEditingMyFields Then Exit Sub
    NeedSave = True
End Sub

' -----------------------------------------------------------------------------
' HoursAndMinets (Function), source lines 79-86
' -----------------------------------------------------------------------------
Private Function HoursAndMinets(min As Long) As String
    Dim h As Long
    Dim m As Long
    h = min \ 60
    m = min Mod 60
    If h > 0 Then HoursAndMinets = h & " " & localization.ActiveLng.Item("ActivationForm.hour") & " "
    If m > 0 Then HoursAndMinets = HoursAndMinets & m & " " & localization.ActiveLng.Item("ActivationForm.minutes") & " "
End Function

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 87-92
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    If NeedSave Then
        UserData.CodeOfActivation = Trim(Text1.text)
        UserData.Save
    End If
End Sub
