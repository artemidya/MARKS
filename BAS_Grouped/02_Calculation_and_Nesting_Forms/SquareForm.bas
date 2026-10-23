' =============================================================================
' Group: 02_Calculation_and_Nesting_Forms
' Source module: SquareForm.frm
' Procedures: 15
' =============================================================================

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 16-60
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    UserData.Load
    FirstRun = True
    CalculateSanP.SquareFormVisible = True

    left = UserData.SquareFormData.WindowPositionX
    CurWindowPositionX = left
    top = UserData.SquareFormData.WindowPositionY
    CurWindowPositionY = top
    
    IsEditingMyFields = True
    
    ComboBox1.AddItem
    ComboBox1.AddItem
    ComboBox1.AddItem
    SetTexts


    If UserData.SquareFormData.RectArea = 1 Then
        ComboBox1.ListIndex = 1
        Image3.Picture = Image2.Picture
    ElseIf UserData.SquareFormData.RectArea = 2 Then
        ComboBox1.ListIndex = 2
        Image3.Picture = Image4.Picture
    Else
        ComboBox1.ListIndex = 0
        Image3.Picture = Image1.Picture
    End If
    

    
    TextBox1 = Format(UserData.SquareFormData.Price / 100, "0.00")
    If CalculateSanP.PerimetrFormVisible = True Then
        TextBox2 = PerimetrForm.TextBox2
    Else
        TextBox2 = Format(UserData.PerimetrFormData.Discount / 100)
    End If
    IsEditingMyFields = False
    

    
    TextBox1.SelStart = 0
    TextBox1.SelLength = 10
    UserFormRefreshData
End Sub

' -----------------------------------------------------------------------------
' SetTexts (Sub), source lines 61-69
' -----------------------------------------------------------------------------
Public Sub SetTexts()
    SquareForm.Caption = localization.ActiveLng.Item("SquareForm.Caption")
    Label3.Caption = localization.ActiveLng.Item("SquareForm.Label3.Caption") & ChrW(178) & ", " & Trim(UserData.PerimetrFormData.MoneyUnit) & "/" & localization.ActiveLng.Item("SquareForm.m") & ChrW(178)
    Label5.Caption = localization.ActiveLng.Item("SquareForm.Label5.Caption")
    Label7.Caption = localization.ActiveLng.Item("SquareForm.Label7.Caption") & ", " & Trim(UserData.PerimetrFormData.MoneyUnit)
    ComboBox1.List(0) = localization.ActiveLng.Item("SquareForm.AreaShapes") & ChrW(178)
    ComboBox1.List(1) = localization.ActiveLng.Item("SquareForm.AreaOfRectangularBars") & ChrW(178)
    ComboBox1.List(2) = localization.ActiveLng.Item("SquareForm.AreaOfSelection") & ChrW(178)
End Sub

' -----------------------------------------------------------------------------
' UserForm_Activate (Sub), source lines 71-82
' -----------------------------------------------------------------------------
Private Sub UserForm_Activate()
    If FirstRun = True Then
        If CalculateSanP.PerimetrFormVisible = False Then
            CalculateSanP.Start
        Else
            Me.Repaint
            If ComboBox1.ListIndex = 1 Then CalculateSanP.ShowRectOnScreenCurves
            If ComboBox1.ListIndex = 2 Then CalculateSanP.SelectionRectOnScreenCurve.Show
        End If
    End If
    FirstRun = False
End Sub

' -----------------------------------------------------------------------------
' ComboBox1_Change (Sub), source lines 83-105
' -----------------------------------------------------------------------------
Private Sub ComboBox1_Change()
    If IsEditingMyFields = False Then
        If ComboBox1.ListIndex = 0 Then
            Image3.Picture = Image1.Picture
            UserFormRefreshData
            Me.Repaint
            CalculateSanP.HideRectOnScreenCurves
            CalculateSanP.SelectionRectOnScreenCurve.Hide
        ElseIf ComboBox1.ListIndex = 1 Then
            Image3.Picture = Image2.Picture
            UserFormRefreshData
            Me.Repaint
            CalculateSanP.ShowRectOnScreenCurves
            CalculateSanP.SelectionRectOnScreenCurve.Hide
        Else
            Image3.Picture = Image4.Picture
            UserFormRefreshData
            Me.Repaint
            CalculateSanP.HideRectOnScreenCurves
            CalculateSanP.SelectionRectOnScreenCurve.Show
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' TextBox1_Change (Sub), source lines 106-108
' -----------------------------------------------------------------------------
Private Sub TextBox1_Change()
    If IsEditingMyFields = False Then UserFormRefreshData
End Sub

' -----------------------------------------------------------------------------
' TextBox2_Change (Sub), source lines 109-119
' -----------------------------------------------------------------------------
Private Sub TextBox2_Change()
    If IsEditingMyFields = False Then
        UserFormRefreshData
        If CalculateSanP.PerimetrFormVisible = True Then
            PerimetrForm.IsEditingMyFields = True
            PerimetrForm.TextBox2 = TextBox2
            PerimetrForm.IsEditingMyFields = False
            PerimetrForm.UserFormRefreshData
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' CheckBox1_KeyDown (Sub), source lines 121-123
' -----------------------------------------------------------------------------
Private Sub CheckBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload SquareForm
End Sub

' -----------------------------------------------------------------------------
' TextBox1_KeyDown (Sub), source lines 124-126
' -----------------------------------------------------------------------------
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload SquareForm
End Sub

' -----------------------------------------------------------------------------
' TextBox2_KeyDown (Sub), source lines 127-129
' -----------------------------------------------------------------------------
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload SquareForm
End Sub

' -----------------------------------------------------------------------------
' TextBox3_KeyDown (Sub), source lines 130-132
' -----------------------------------------------------------------------------
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload SquareForm
End Sub

' -----------------------------------------------------------------------------
' TextBox5_KeyDown (Sub), source lines 133-135
' -----------------------------------------------------------------------------
Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload SquareForm
End Sub

' -----------------------------------------------------------------------------
' UserForm_Layout (Sub), source lines 136-139
' -----------------------------------------------------------------------------
Private Sub UserForm_Layout()
    CurWindowPositionX = left
    CurWindowPositionY = top
End Sub

' -----------------------------------------------------------------------------
' UserFormRefreshData (Sub), source lines 141-177
' -----------------------------------------------------------------------------
Public Sub UserFormRefreshData()
    On Error Resume Next
    Dim k As Double, errStr As String
    Dim Area As Double
    
    If CountErr > 0 And ComboBox1.ListIndex <> 2 Then
        errStr = "err(" & CalculateSanP.CountErr & ") "
        TextBox3.ForeColor = RGB(250, 0, 0)
        TextBox3.ControlTipText = localization.ActiveLng.Item("SquareForm.TextBox3.ControlTipText") & " (" & CalculateSanP.CountErr & " " & localization.ActiveLng.Item("SquareForm.pcs") & ")"
    Else
        errStr = ""
        TextBox3.ForeColor = TextBox2.ForeColor
        TextBox3.ControlTipText = ""
    End If

    Dim k2 As Double: k2 = VGCore.Application.ConvertUnits(1, CalculateSanP.lastUnit, cdrInch)
    k = 0.00064516 * k2 * k2
    
    If ComboBox1.ListIndex = 1 Then
        TextBox3 = errStr & CalculateSanP.ProcessingStr & Format(CalculateSanP.RectAreaOfCurves * k, "0.000000")
        TextBox5.text = CalculateSanP.ProcessingStr & Format(CalculateSanP.RectAreaOfCurves * k * ValP(TextBox1.text) * (100 + ValP(TextBox2)) / 100, "0.00")
        CalculateSanP.ErrShow
    ElseIf ComboBox1.ListIndex = 2 Then
        TextBox3 = Format(CalculateSanP.SelectionRectArea * k, "0.000000")
        TextBox5.text = Format(CalculateSanP.SelectionRectArea * k * ValP(TextBox1.text) * (100 + ValP(TextBox2)) / 100, "0.00")
        If CalculateSanP.PerimetrFormVisible Then
            CalculateSanP.ErrShow
        Else
            CalculateSanP.ErrHide
        End If
    Else
        TextBox3 = errStr & CalculateSanP.ProcessingStr & Format(CalculateSanP.AreaOfCurves * k, "0.000000")
        TextBox5.text = CalculateSanP.ProcessingStr & Format(CalculateSanP.AreaOfCurves * k * ValP(TextBox1.text) * (100 + ValP(TextBox2)) / 100, "0.00")
        CalculateSanP.ErrShow
    End If
    
End Sub

' -----------------------------------------------------------------------------
' ValP (Function), source lines 178-181
' -----------------------------------------------------------------------------
Private Function ValP(str As String) As Double
    str = Replace(str, ",", ".")
    ValP = val(str)
End Function

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 183-204
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    On Error Resume Next
    If CalculateSanP.PerimetrFormVisible = False Then
        CalculateSanP.Break
    End If
    CalculateSanP.SquareFormVisible = False
    CalculateSanP.HideRectOnScreenCurves
    CalculateSanP.SelectionRectOnScreenCurve.Hide
    
    If UserData.SquareFormData.Price <> ValP(TextBox1.text) * 100 _
    Or UserData.PerimetrFormData.Discount <> ValP(TextBox2.text) * 100 _
    Or UserData.SquareFormData.WindowPositionX <> CurWindowPositionX _
    Or UserData.SquareFormData.WindowPositionY <> CurWindowPositionY _
    Or UserData.SquareFormData.RectArea <> ComboBox1.ListIndex Then
        UserData.SquareFormData.Price = ValP(TextBox1.text) * 100
        UserData.PerimetrFormData.Discount = ValP(TextBox2.text) * 100
        UserData.SquareFormData.WindowPositionX = CurWindowPositionX
        UserData.SquareFormData.WindowPositionY = CurWindowPositionY
        UserData.SquareFormData.RectArea = ComboBox1.ListIndex
        UserData.Save
    End If
End Sub
