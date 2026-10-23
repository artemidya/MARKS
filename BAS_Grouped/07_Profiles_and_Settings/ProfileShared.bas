' =============================================================================
' Group: 07_Profiles_and_Settings
' Source module: ProfileShared.bas
' Procedures: 12
' =============================================================================

' -----------------------------------------------------------------------------
' ListIndexByName (Function), source lines 33-48
' -----------------------------------------------------------------------------
Public Function ListIndexByName(name As String, Combo As ComboBox, Optional defIndex As Long = 0) As Long
    Dim q As Long
    With Combo
        For q = 0 To .ListCount - 1
            If KillBirdsAndSpaces(.List(q)) = name Then
                ListIndexByName = q
                Exit Function
            End If
        Next
        If defIndex < .ListCount Then
            ListIndexByName = defIndex 'Select deafault
        Else
            ListIndexByName = 0
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' StringToDouble (Function), source lines 49-53
' -----------------------------------------------------------------------------
Public Function StringToDouble(text As String) As Double
    On Error GoTo 10
    StringToDouble = val(text)
10
End Function

' -----------------------------------------------------------------------------
' StringToLong (Function), source lines 54-62
' -----------------------------------------------------------------------------
Public Function StringToLong(text As String) As Long
    On Error GoTo 10
    Dim V As Double
    V = val(text)
    If V > 2147483647 Then V = 2147483647
    If V < -2147483646 Then V = -2147483646
    StringToLong = CLng(V)
10
End Function

' -----------------------------------------------------------------------------
' LongArrayToString (Function), source lines 63-71
' -----------------------------------------------------------------------------
Public Function LongArrayToString(arr() As Long) As String
    If Not Not arr Then
        Dim q As Long
        For q = 0 To UBound(arr) - 1
            LongArrayToString = LongArrayToString & arr(q) & ";"
        Next
        LongArrayToString = LongArrayToString & arr(q)
    End If
End Function

' -----------------------------------------------------------------------------
' StringToLongArray (Function), source lines 72-85
' -----------------------------------------------------------------------------
Public Function StringToLongArray(text As String) As Long()
    text = Trim(text)
    Dim arrStr() As String
    arrStr = Split(text, ";")
    If UBound(arrStr) > -1 Then
        Dim arr() As Long
        ReDim arr(UBound(arrStr))
        Dim q As Long
        For q = 0 To UBound(arrStr)
            arr(q) = StringToLong(arrStr(q))
        Next
    End If
    StringToLongArray = arr
End Function

' -----------------------------------------------------------------------------
' BooleanToString (Function), source lines 86-92
' -----------------------------------------------------------------------------
Public Function BooleanToString(value As Boolean) As String
    If value Then
        BooleanToString = "1"
    Else
        BooleanToString = "0"
    End If
End Function

' -----------------------------------------------------------------------------
' StringToBoolean (Function), source lines 93-95
' -----------------------------------------------------------------------------
Public Function StringToBoolean(text As String) As Boolean
    If text = "1" Then StringToBoolean = True
End Function

' -----------------------------------------------------------------------------
' CommandLanguageToString (Function), source lines 96-105
' -----------------------------------------------------------------------------
Public Function CommandLanguageToString(cl As CommandLanguages) As String
    Select Case cl
        Case HPGL
            CommandLanguageToString = "HPGL"
        Case DMPL
            CommandLanguageToString = "DMPL"
        Case Else
            CommandLanguageToString = "HPGL"
    End Select
End Function

' -----------------------------------------------------------------------------
' StringToCommandLanguage (Function), source lines 106-116
' -----------------------------------------------------------------------------
Public Function StringToCommandLanguage(text As String) As CommandLanguages
    text = UCase(text)
    Select Case text
        Case "HPGL"
            StringToCommandLanguage = HPGL
        Case "DMPL"
            StringToCommandLanguage = DMPL
        Case Else
            StringToCommandLanguage = HPGL
    End Select
End Function

' -----------------------------------------------------------------------------
' StringToCoordinateSystem (Function), source lines 117-127
' -----------------------------------------------------------------------------
Public Function StringToCoordinateSystem(text As String) As CoordinateSystems
    text = UCase(text)
    Select Case text
        Case "ABSOLUTE"
            StringToCoordinateSystem = Absolute
        Case "RELATIVE"
            StringToCoordinateSystem = Relative
        Case Else
            StringToCoordinateSystem = Absolute
    End Select
End Function

' -----------------------------------------------------------------------------
' CoordinateSystemToString (Function), source lines 128-137
' -----------------------------------------------------------------------------
Public Function CoordinateSystemToString(cs As CoordinateSystems) As String
    Select Case cs
        Case Absolute
            CoordinateSystemToString = "ABSOLUTE"
        Case Relative
            CoordinateSystemToString = "RELATIVE"
        Case Else
            CoordinateSystemToString = "ABSOLUTE"
    End Select
End Function

' -----------------------------------------------------------------------------
' KillBirdsAndSpaces (Function), source lines 138-157
' -----------------------------------------------------------------------------
Private Function KillBirdsAndSpaces(ByVal s As String) As String
    ' Удалить 4 пробела в начале, если строка достаточно длинная
    If Len(s) >= 4 Then
        If left(s, 4) = "    " Then
            KillBirdsAndSpaces = Mid(s, 5)
            Exit Function
        End If
    End If

    ' Удалить символ ChrW(10003) (?) в начале, если строка не пустая
    If Len(s) >= 1 Then
        If left(s, 1) = ChrW(10003) Then
            KillBirdsAndSpaces = Mid(s, 2)
            Exit Function
        End If
    End If

    ' Вернуть как есть, если ничего не найдено
    KillBirdsAndSpaces = s
End Function
