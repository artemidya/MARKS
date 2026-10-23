' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: ConverterStringLToPlotCalcCrv.cls
' Procedures: 11
' =============================================================================

' -----------------------------------------------------------------------------
' GetPlotCalcCrv (Function), source lines 39-60
' -----------------------------------------------------------------------------
Public Function GetPlotCalcCrv(ByRef StringL As String) As Collection
    On Error GoTo 10
    Set GetPlotCalcCrv = New Collection
    LenStringL = Len(StringL)
    cp = 1
    Dim brush As waBrush
    Dim path As waPath
    Do While cp <= LenStringL
        '1 GetBrush
        Set brush = GetBrush(StringL)
        '2 GetPath
        Set path = GetPath(StringL)
        '3 Add to Collection
        
        If Not path Is Nothing Then
            Set path.brush = brush
            GetPlotCalcCrv.Add path
        End If
        cp = cp + 1
    Loop
10
End Function

' -----------------------------------------------------------------------------
' GetPath (Function), source lines 61-123
' -----------------------------------------------------------------------------
Private Function GetPath(ByRef StringL As String) As waPath
    Static bstr As String
    Dim isPathCreated As Boolean
    Dim IsSubpathStarted As Boolean
    Dim LastX As Long
    Dim LastY As Long
    
    Dim x2 As Long
    Dim y2 As Long
    Dim x3 As Long
    Dim y3 As Long
    Dim x4 As Long
    Dim y4 As Long
    
    Set GetPath = New waPath

    Do While cp <= LenStringL
        If IsPath(StringL) Then
            bstr = Mid(StringL, cp, 1)
            If bstr = "U" Then
                'New subpath
                If IsSubpathStarted Then
                    GetPath.CloseSubPath
                End If

                isPathCreated = True
                GetPath.StartSubPath
                LastX = GetCoordinate(StringL)
                LastY = GetCoordinate(StringL)
                GetPath.AddLineSegment LastX, LastY
                IsSubpathStarted = True
                
            ElseIf bstr = "D" Then
                'Add line
                If IsSubpathStarted Then
                    LastX = GetCoordinate(StringL)
                    LastY = GetCoordinate(StringL)
                    GetPath.AddLineSegment LastX, LastY
                End If
            ElseIf bstr = "B" Then
                'Add bezier
                If IsSubpathStarted Then
                    x2 = GetCoordinate(StringL)
                    y2 = GetCoordinate(StringL)
                    x3 = GetCoordinate(StringL)
                    y3 = GetCoordinate(StringL)
                    x4 = GetCoordinate(StringL)
                    y4 = GetCoordinate(StringL)
                    GetPath.AddBezierSegment LastX, LastY, x2, y2, x3, y3, x4, y4
                    LastX = x4
                    LastY = y4
                End If
            End If
        Else
            cp = cp - 1
            Exit Do
        End If
    Loop
'    GetPath.CloseSubPath
    If Not isPathCreated Then
        Set GetPath = Nothing
    End If
End Function

' -----------------------------------------------------------------------------
' GetBrush (Function), source lines 124-144
' -----------------------------------------------------------------------------
Private Function GetBrush(ByRef StringL As String) As waBrush
    For cp = cp To LenStringL
        If IsBrush(StringL) Then
            If Mid(StringL, cp, 1) = "S" Then
                'Solid
                Dim col As RGBA
                col = GetSolidBrushColor(StringL)
                Set GetBrush = CreateBrush(RGB(col.R, col.g, col.b), CLng(CDbl(col.a) / 255 * 100))
            Else
                'Gradient
                Dim gbp As GradientBrushParametrs
                gbp = GetGradientBrushColors(StringL)
                Set GetBrush = New waBrush
                With gbp
                    GetBrush.SetLinearGradient .x1, .y1, .x2, .y2, RGB(.Color1.R, .Color1.g, .Color1.b), RGB(.Color2.R, .Color2.g, .Color2.b), , CLng(CDbl(.Color1.a) / 255 * 100), CLng(CDbl(.Color2.a) / 255 * 100)
                End With
            End If
            Exit Function
        End If
    Next
End Function

' -----------------------------------------------------------------------------
' GetGradientBrushColors (Function), source lines 145-152
' -----------------------------------------------------------------------------
Private Function GetGradientBrushColors(ByRef StringL As String) As GradientBrushParametrs
    GetGradientBrushColors.Color1 = GetSolidBrushColor(StringL)
    GetGradientBrushColors.Color2 = GetSolidBrushColor(StringL)
    GetGradientBrushColors.x1 = GetCoordinate(StringL)
    GetGradientBrushColors.y1 = GetCoordinate(StringL)
    GetGradientBrushColors.x2 = GetCoordinate(StringL)
    GetGradientBrushColors.y2 = GetCoordinate(StringL)
End Function

' -----------------------------------------------------------------------------
' GetCoordinate (Function), source lines 153-162
' -----------------------------------------------------------------------------
Private Function GetCoordinate(ByRef StringL As String) As Long
    Dim d As String
    
    cp = cp + 1
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetCoordinate = TxtToLong(d)
End Function

' -----------------------------------------------------------------------------
' GetSolidBrushColor (Function), source lines 163-198
' -----------------------------------------------------------------------------
Private Function GetSolidBrushColor(ByRef StringL As String) As RGBA
    Dim d As String
    GetSolidBrushColor.a = 255
    
    cp = cp + 1
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetSolidBrushColor.R = TxtToByte(d)
    
    cp = cp + 1
    d = ""
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetSolidBrushColor.g = TxtToByte(d)
    
    cp = cp + 1
    d = ""
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetSolidBrushColor.b = TxtToByte(d)
    
    cp = cp + 1
    d = ""
    Do While cp <= LenStringL
        If IsDigit(StringL) Then d = d & Mid(StringL, cp, 1) Else Exit Do
        cp = cp + 1
    Loop
    GetSolidBrushColor.a = TxtToByte(d)
    
End Function

' -----------------------------------------------------------------------------
' TxtToByte (Function), source lines 199-202
' -----------------------------------------------------------------------------
Private Function TxtToByte(ByRef txt As String) As Byte
    On Error Resume Next
    TxtToByte = CByte(txt)
End Function

' -----------------------------------------------------------------------------
' TxtToLong (Function), source lines 203-206
' -----------------------------------------------------------------------------
Private Function TxtToLong(ByRef txt As String) As Long
    On Error Resume Next
    TxtToLong = CLng(txt)
End Function

' -----------------------------------------------------------------------------
' IsPath (Function), source lines 208-212
' -----------------------------------------------------------------------------
Private Function IsPath(ByRef StringL As String) As Boolean
    Static bstr As String
    bstr = Mid(StringL, cp, 1)
    If bstr = "U" Or bstr = "D" Or bstr = "B" Then IsPath = True
End Function

' -----------------------------------------------------------------------------
' IsBrush (Function), source lines 213-218
' -----------------------------------------------------------------------------
Private Function IsBrush(ByRef StringL As String) As Boolean
    Static bstr As String
    bstr = Mid(StringL, cp, 1)
    If bstr = "S" Then IsBrush = True: Exit Function
    If bstr = "G" Then IsBrush = True
End Function

' -----------------------------------------------------------------------------
' IsDigit (Function), source lines 219-224
' -----------------------------------------------------------------------------
Private Function IsDigit(ByRef StringL As String) As Boolean
    '[48-57] - Digits "-" - 45
    Static ascii As Long
    ascii = Asc(Mid(StringL, cp, 1))
    If (ascii > 47 And ascii < 58) Or ascii = 45 Then IsDigit = True
End Function
