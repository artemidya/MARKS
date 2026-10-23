' =============================================================================
' Group: 02_Calculation_and_Nesting_Forms
' Source module: PerimetrForm.frm
' Procedures: 13
' =============================================================================

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 17-41
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    UserData.Load
    FirstRun = True
    CalculateSanP.PerimetrFormVisible = True
    
    left = UserData.PerimetrFormData.WindowPositionX
    CurWindowPositionX = left
    top = UserData.PerimetrFormData.WindowPositionY
    CurWindowPositionY = top
    SetTexts
    
    IsEditingMyFields = True
    TextBox1 = Format(UserData.PerimetrFormData.Price / 100, "0.00")
    If CalculateSanP.SquareFormVisible = True Then
        TextBox2 = SquareForm.TextBox2
    Else
        TextBox2 = Format(UserData.PerimetrFormData.Discount / 100)
    End If
    IsEditingMyFields = False
    
    TextBox1.SelStart = 0
    TextBox1.SelLength = 10
    CalculateSanP.ErrShow
    UserFormRefreshData
End Sub

' -----------------------------------------------------------------------------
' SetTexts (Sub), source lines 42-51
' -----------------------------------------------------------------------------
Public Sub SetTexts()
    PerimetrForm.Caption = localization.ActiveLng.Item("PerimetrForm.Caption")
    Label1.Caption = localization.ActiveLng.Item("PerimetrForm.Label1.Caption")
    Label8.Caption = localization.ActiveLng.Item("PerimetrForm.Label8.Caption")
    Label9.Caption = localization.ActiveLng.Item("PerimetrForm.Label9.Caption")
    
    Label3.Caption = localization.ActiveLng.Item("PerimetrForm.Label3.Caption") & ", " & Trim(UserData.PerimetrFormData.MoneyUnit) & "/" & localization.ActiveLng.Item("PerimetrForm.m")
    Label5.Caption = localization.ActiveLng.Item("PerimetrForm.Label5.Caption")
    Label7.Caption = localization.ActiveLng.Item("PerimetrForm.Label7.Caption") & ", " & Trim(UserData.PerimetrFormData.MoneyUnit)
End Sub

' -----------------------------------------------------------------------------
' UserForm_Activate (Sub), source lines 53-60
' -----------------------------------------------------------------------------
Private Sub UserForm_Activate()
    If CalculateSanP.SquareFormVisible = False Then
       If FirstRun = True Then
            CalculateSanP.Start
       End If
    End If
    FirstRun = False
End Sub

' -----------------------------------------------------------------------------
' TextBox1_Change (Sub), source lines 61-63
' -----------------------------------------------------------------------------
Private Sub TextBox1_Change()
    If IsEditingMyFields = False Then UserFormRefreshData
End Sub

' -----------------------------------------------------------------------------
' TextBox2_Change (Sub), source lines 64-74
' -----------------------------------------------------------------------------
Private Sub TextBox2_Change()
    If IsEditingMyFields = False Then
        UserFormRefreshData
        If CalculateSanP.SquareFormVisible = True Then
            SquareForm.IsEditingMyFields = True
            SquareForm.TextBox2 = TextBox2
            SquareForm.IsEditingMyFields = False
            SquareForm.UserFormRefreshData
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' TextBox1_KeyDown (Sub), source lines 75-77
' -----------------------------------------------------------------------------
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload PerimetrForm
End Sub

' -----------------------------------------------------------------------------
' TextBox2_KeyDown (Sub), source lines 78-80
' -----------------------------------------------------------------------------
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload PerimetrForm
End Sub

' -----------------------------------------------------------------------------
' TextBox3_KeyDown (Sub), source lines 81-83
' -----------------------------------------------------------------------------
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload PerimetrForm
End Sub

' -----------------------------------------------------------------------------
' TextBox5_KeyDown (Sub), source lines 84-86
' -----------------------------------------------------------------------------
Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload PerimetrForm
End Sub

' -----------------------------------------------------------------------------
' UserForm_Layout (Sub), source lines 89-92
' -----------------------------------------------------------------------------
Private Sub UserForm_Layout()
    CurWindowPositionX = left
    CurWindowPositionY = top
End Sub

' -----------------------------------------------------------------------------
' UserFormRefreshData (Sub), source lines 94-113
' -----------------------------------------------------------------------------
Public Sub UserFormRefreshData()
    On Error Resume Next
    Dim k As Double, errStr As String
    If CountErr > 0 Then
        errStr = "err(" & CalculateSanP.CountErr & ") "
        TextBox3.ForeColor = RGB(250, 0, 0)
        TextBox3.ControlTipText = localization.ActiveLng.Item("PerimetrForm.ImpossibleMeasureLengthOfCurve") & " (" & CalculateSanP.CountErr & " " & localization.ActiveLng.Item("PerimetrForm.unit") & ")"
        CalculateSanP.ErrShow
    Else
        errStr = ""
        TextBox3.ForeColor = TextBox2.ForeColor
        TextBox3.ControlTipText = ""
    End If
    
    k = 0.0254 * VGCore.Application.ConvertUnits(1, CalculateSanP.lastUnit, cdrInch)
    TextBox3.text = errStr & CalculateSanP.ProcessingStr & Format(CalculateSanP.LengthOfCurves * k, "0.000")
    TextBox5.text = CalculateSanP.ProcessingStr & Format(CalculateSanP.LengthOfCurves * k * ValP(TextBox1.text) * (100 + ValP(TextBox2)) / 100, "0.00")
    TextBox6.text = CalculateSanP.ProcessingStr & CalculateSanP.CountSubPaths
    TextBox7.text = CalculateSanP.ProcessingStr & CalculateSanP.CountNodes
End Sub

' -----------------------------------------------------------------------------
' ValP (Function), source lines 114-117
' -----------------------------------------------------------------------------
Private Function ValP(str As String) As Double
    str = Replace(str, ",", ".")
    ValP = val(str)
End Function

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 118-137
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    On Error Resume Next
    If CalculateSanP.SquareFormVisible = False Then
       CalculateSanP.Break
    Else
        If SquareForm.ComboBox1.ListIndex = 2 Then CalculateSanP.ErrHide
    End If
    CalculateSanP.PerimetrFormVisible = False
    
    If UserData.PerimetrFormData.Price <> ValP(TextBox1.text) * 100 _
    Or UserData.PerimetrFormData.Discount <> ValP(TextBox2.text) * 100 _
    Or UserData.PerimetrFormData.WindowPositionX <> CurWindowPositionX _
    Or UserData.PerimetrFormData.WindowPositionY <> CurWindowPositionY Then
        UserData.PerimetrFormData.Price = ValP(TextBox1.text) * 100
        UserData.PerimetrFormData.Discount = ValP(TextBox2.text) * 100
        UserData.PerimetrFormData.WindowPositionX = CurWindowPositionX
        UserData.PerimetrFormData.WindowPositionY = CurWindowPositionY
        UserData.Save
    End If
End Sub
