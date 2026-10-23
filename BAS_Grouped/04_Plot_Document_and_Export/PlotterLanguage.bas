' =============================================================================
' Group: 04_Plot_Document_and_Export
' Source module: PlotterLanguage.bas
' Procedures: 12
' =============================================================================

' -----------------------------------------------------------------------------
' Init (Function), source lines 21-28
' -----------------------------------------------------------------------------
Public Function Init() As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            Init = "IN;"
        Case DMPL
            Init = ";:HL0,EC1,UP1,"
    End Select
End Function

' -----------------------------------------------------------------------------
' TERM (Function), source lines 29-36
' -----------------------------------------------------------------------------
Public Function TERM() As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            TERM = ""
        Case DMPL
            TERM = "U@"
    End Select
End Function

' -----------------------------------------------------------------------------
' SetPA (Function), source lines 38-45
' -----------------------------------------------------------------------------
Public Function SetPA() As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetPA = "PA"
        Case DMPL
            SetPA = "A"
    End Select
End Function

' -----------------------------------------------------------------------------
' SetPR (Function), source lines 46-53
' -----------------------------------------------------------------------------
Public Function SetPR() As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetPR = "PR"
        Case DMPL
            SetPR = "R"
    End Select
End Function

' -----------------------------------------------------------------------------
' SetAS (Function), source lines 54-67
' -----------------------------------------------------------------------------
Public Function SetAS(acceleration As Long) As String
    If acceleration = 0 Then Exit Function
    'as = 1-4
    ' 1 -> 1,25 m/s^2
    ' 2 -> 2,5 m/s^2
    ' 3 -> 5 m/s^2
    ' 4 -> 10 m/s^2
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetAS = "AS" & CStr(acceleration)
        Case DMPL
            'Exit (no acceleration command in DMPL)
    End Select
End Function

' -----------------------------------------------------------------------------
' SetVS (Function), source lines 68-77
' -----------------------------------------------------------------------------
Public Function SetVS(velocity As Long) As String
    velocity = velocity \ 10
    If velocity = 0 Then Exit Function
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetVS = "VS" & CStr(velocity)
        Case DMPL
            SetVS = "V" & CStr(velocity)
    End Select
End Function

' -----------------------------------------------------------------------------
' SetFS (Function), source lines 78-86
' -----------------------------------------------------------------------------
Public Function SetFS(force As Long) As String
    If force = 0 Then Exit Function
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetFS = "FS" & CStr(force)
        Case DMPL
            'Exit (no force command in DMPL)
    End Select
End Function

' -----------------------------------------------------------------------------
' MoveCursor (Function), source lines 87-96
' -----------------------------------------------------------------------------
Public Function MoveCursor(x As Long, y As Long, Optional velocity As Long = 50) As String
    velocity = velocity \ 10
    If velocity < 1 Then velocity = 1
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            MoveCursor = "PA;VS" & velocity & ";PU" & CStr(x) & "," & CStr(y) & ";"
        Case DMPL
            MoveCursor = "AV" & velocity & ",U" & CStr(x) & "," & CStr(y) & ","
    End Select
End Function

' -----------------------------------------------------------------------------
' MovePU (Function), source lines 97-104
' -----------------------------------------------------------------------------
Public Function MovePU(x As Long, y As Long) As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            MovePU = "PU" & CStr(x) & "," & CStr(y)
        Case DMPL
            MovePU = "U" & CStr(x) & "," & CStr(y) & ","
    End Select
End Function

' -----------------------------------------------------------------------------
' MovePD (Function), source lines 105-112
' -----------------------------------------------------------------------------
Public Function MovePD(x As Long, y As Long) As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            MovePD = "PD" & CStr(x) & "," & CStr(y)
        Case DMPL
            MovePD = "D" & CStr(x) & "," & CStr(y) & ","
    End Select
End Function

' -----------------------------------------------------------------------------
' StrToAscii (Function), source lines 116-127
' -----------------------------------------------------------------------------
Public Function StrToAscii(str As String) As Byte()
    Dim l As Long: l = Len(str)
    If l > 0 Then
        Dim bytes() As Byte
        ReDim bytes(l - 1)
        Dim q As Long
        For q = 0 To l - 1
            bytes(q) = Asc(Mid(str, q + 1, 1))
        Next
        StrToAscii = bytes
    End If
End Function

' -----------------------------------------------------------------------------
' AsciiToString (Function), source lines 128-135
' -----------------------------------------------------------------------------
Public Function AsciiToString(bytes() As Byte) As String
    If Not Not bytes Then
        Dim q As Long
        For q = 0 To UBound(bytes)
            AsciiToString = AsciiToString & Chr(bytes(q))
        Next
    End If
End Function
