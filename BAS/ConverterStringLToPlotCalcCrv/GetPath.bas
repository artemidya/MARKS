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
