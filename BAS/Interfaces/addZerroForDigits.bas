Private Function addZerroForDigits(ByVal txt As String, Optional ByVal maxDigits As Long = 50) As String
    txt = Replace(txt, ",", ".")
    Dim q As Long
    Dim code As Integer
    Dim lastCode As Integer
    Dim isStartingDigit As Boolean
    Dim startPositionDigits() As Long
    Dim endPositionDigits() As Long
    Dim digitsCount As Long
    
    Dim toRight() As Boolean ' (???? ?????, ?? ???????? ???? ??????)
    
    For q = 1 To Len(txt)
        code = Asc(Mid(txt, q, 1))
        If (code >= 48 And code <= 57) Then '..  Or code = 46
            If isStartingDigit = False Then
                isStartingDigit = True
                digitsCount = digitsCount + 1
                ReDim Preserve startPositionDigits(digitsCount - 1)
                ReDim Preserve endPositionDigits(digitsCount - 1)
                ReDim Preserve toRight(digitsCount - 1)
                startPositionDigits(digitsCount - 1) = q
                If lastCode = 46 Then toRight(digitsCount - 1) = True
            End If
        Else
            If isStartingDigit = True Then
                isStartingDigit = False
                endPositionDigits(digitsCount - 1) = q
            End If
        End If
        lastCode = code
    Next
    
    If isStartingDigit = True Then endPositionDigits(digitsCount - 1) = q
    
    Dim s As Long: s = 1
    Dim l As Long
    Dim l2 As Long
    Dim newDigits As String
    Dim q2 As Long

    For q = 0 To digitsCount - 1
        l = startPositionDigits(q) - s
        addZerroForDigits = addZerroForDigits & Mid(txt, s, l)
        l2 = endPositionDigits(q) - startPositionDigits(q)
        s = s + l + l2
        
        newDigits = ""
        For q2 = 1 To maxDigits - l2
            newDigits = newDigits & "0"
        Next
        If toRight(q) = True Then
            newDigits = Mid(txt, startPositionDigits(q), l2) & newDigits
        Else
            newDigits = newDigits & Mid(txt, startPositionDigits(q), l2)
        End If
        addZerroForDigits = addZerroForDigits & newDigits
    Next
    l = Len(txt) - s + 1
    addZerroForDigits = addZerroForDigits & Mid(txt, s, l)
End Function
