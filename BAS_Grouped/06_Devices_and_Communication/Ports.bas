' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: Ports.cls
' Procedures: 22
' =============================================================================

' -----------------------------------------------------------------------------
' PortType (PropertyGet), source lines 63-73
' -----------------------------------------------------------------------------
Property Get PortType(OutPutName As String) As PortType
    Dim i As Long
    i = GetIndex(OutPutName)
    If UCase(left(Ports(i).name, 3)) = "COM" Then
        PortType = ComPort
    ElseIf UCase(left(Ports(i).name, 3)) = "LPT" Then
        PortType = LPTport
    Else
        PortType = UnknownPort
    End If
End Property

' -----------------------------------------------------------------------------
' Connected (PropertyGet), source lines 75-77
' -----------------------------------------------------------------------------
Property Get Connected(ByVal i As Long) As Boolean
    Connected = Ports(i).Active
End Property

' -----------------------------------------------------------------------------
' count (PropertyGet), source lines 78-80
' -----------------------------------------------------------------------------
Property Get count() As Long
    count = lCount
End Property

' -----------------------------------------------------------------------------
' Names (PropertyGet), source lines 81-83
' -----------------------------------------------------------------------------
Property Get Names(i As Long) As String
    Names = Ports(i).name
End Property

' -----------------------------------------------------------------------------
' Descriptions (PropertyGet), source lines 84-86
' -----------------------------------------------------------------------------
Property Get Descriptions(i As Long) As String
    Descriptions = Ports(i).Description
End Property

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 88-90
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Refresh
End Sub

' -----------------------------------------------------------------------------
' x (Function), source lines 92-96
' -----------------------------------------------------------------------------
Public Function x(OutPutName As String) As Long
    Dim i As Long
    i = GetIndex(OutPutName)
    x = Ports(i).x
End Function

' -----------------------------------------------------------------------------
' y (Function), source lines 97-101
' -----------------------------------------------------------------------------
Public Function y(OutPutName As String) As Long
    Dim i As Long
    i = GetIndex(OutPutName)
    y = Ports(i).y
End Function

' -----------------------------------------------------------------------------
' ComSettingsStr (PropertyGet), source lines 102-112
' -----------------------------------------------------------------------------
Property Get ComSettingsStr(OutPutName As String) As String
    On Error GoTo 10
    Dim i As Long:    i = GetIndex(OutPutName)
    Dim RegData1 As New RegData
    With Ports(i)
        If InStr(1, .name, "COM", vbTextCompare) Then
            ComSettingsStr = RegData1.ComSettingsStr(RemoveTrailingColon(Trim(.name)))
        End If
    End With
10
End Property

' -----------------------------------------------------------------------------
' Out (Sub), source lines 113-158
' -----------------------------------------------------------------------------
Public Sub Out(OutPutName As String, ByRef bytes() As Byte, x As Long, y As Long, Optional dntInitialize As Boolean = False)
'    COM.PortIsWorking = True
    On Error GoTo 10
    Dim i As Long:    i = GetIndex(OutPutName)
    With Ports(i)
        If (.Init <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer _
        Or .TERM <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator) And Not dntInitialize Then
            If .Initialized Then DeInitialize (i)
            Initialize (i)
        End If
        .x = x
        .y = y

        If UBound(bytes) > -1 Then
            If InStr(1, .name, "COM", vbTextCompare) Then
                COM.CommOpen 1, Trim(.name), RemoveTrailingColon(Trim(.name))
                COM.CommClose 1
            End If
            
            Dim fileNumber As Integer
            fileNumber = FreeFile
                      
            Open Trim(.name) For Binary Access Read Write As #fileNumber
                Dim q As Long, ostatok As Long
                Dim size As Long: size = 16
                Dim outBytes() As Byte
                ReDim outBytes(size - 1)
                ostatok = (UBound(bytes) + 1) Mod size
                For q = 1 To (UBound(bytes) + 1) \ size
                    MoveMemory outBytes(0), bytes((q - 1) * size), size
                    Put #fileNumber, , outBytes
                    DoEvents
                Next
                'send ostatok
                If ostatok > 0 Then
                    ReDim outBytes(ostatok - 1)
                    MoveMemory outBytes(0), bytes((q - 1) * size), ostatok
                    Put #fileNumber, , outBytes
                End If

            Close #fileNumber
'            COM.PortIsWorking = False
        End If
    End With
10
End Sub

' -----------------------------------------------------------------------------
' Initialize (Sub), source lines 159-181
' -----------------------------------------------------------------------------
Private Sub Initialize(i As Long)
    On Error GoTo 10
    With Ports(i)
'        .profileIndex = PlotterForm.TabStrip1.SelectedItem.index
        .Init = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer
        .TERM = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator
        .Initialized = True
        
        PlotterForm.AddLog .Init
        Dim bytes() As Byte
        bytes = StrToAscii(.Init)
        
        If UBound(bytes) > -1 Then
            COM.CommOpen 1, Trim(.name), RemoveTrailingColon(Trim(.name))
            COM.CommClose 1
            Dim fileNumber As Integer: fileNumber = FreeFile
            Open Trim(.name) For Binary Access Read Write As #fileNumber
                Put #fileNumber, , bytes
            Close #fileNumber
        End If
    End With
10
End Sub

' -----------------------------------------------------------------------------
' DeInitialize (Sub), source lines 182-203
' -----------------------------------------------------------------------------
Private Sub DeInitialize(i As Long)
    On Error GoTo 10
    With Ports(i)
        PlotterForm.AddLog .TERM
        Dim bytes() As Byte
        bytes = StrToAscii(.TERM)
        
        .Initialized = False
        .Init = ""
        .TERM = ""

        If UBound(bytes) > -1 Then
            COM.CommOpen 1, Trim(.name), RemoveTrailingColon(Trim(.name))
            COM.CommClose 1
            Dim fileNumber As Integer: fileNumber = FreeFile
            Open Trim(.name) For Binary Access Read Write As #fileNumber
                Put #fileNumber, , bytes
            Close #fileNumber
        End If
    End With
10
End Sub

' -----------------------------------------------------------------------------
' GetIndex (Function), source lines 204-212
' -----------------------------------------------------------------------------
Private Function GetIndex(OutPutName As String) As Long
    Dim q As Long
    For q = 0 To UBound(Ports)
        If Ports(q).name = OutPutName Then
            GetIndex = q
            Exit Function
        End If
    Next
End Function

' -----------------------------------------------------------------------------
' Refresh (Sub), source lines 213-259
' -----------------------------------------------------------------------------
Public Sub Refresh()
    Class_Terminate 'DeInitialize
    Erase Ports()
    lCount = 0
    On Error GoTo 10
    Dim bufferSize As Long
    Dim buffer() As Byte
    Dim lPorts() As PORT_INFO_2
    Dim cPort As deviceInfo

    ' Determine number of ports
    EnumPorts vbNullChar, 2, ByVal 0&, 0&, bufferSize, lCount
    ' Create buffer and get port info
    ReDim buffer(bufferSize - 1)
    EnumPorts vbNullChar, 2, buffer(0), bufferSize, bufferSize, lCount
    ReDim lPorts(lCount - 1)

    MoveMemory lPorts(0), buffer(0), LenB(lPorts(0)) * lCount
    ' Get the name of each port
    Dim q As Long
    For q = 0 To lCount - 1
        With cPort
            .name = PtrToString(lPorts(q).pPortName)
            .Description = PtrToString(lPorts(q).pDescription)
            .MonitorName = PtrToString(lPorts(q).pMonitorName)
            .Active = False
           
            If InStr(1, UCase(.name), "COM") Then
                .Active = IsCOMPortAvailable(.name)
                AddToArray Ports(), cPort
            ElseIf InStr(1, UCase(.name), "LPT") Then
                .Active = True
                AddToArray Ports(), cPort
            End If
            
        End With
        
    Next
    
    If Not Not Ports() Then
        lCount = UBound(Ports) + 1
    Else
        lCount = 0
    End If
    Sort Ports()
10
End Sub

' -----------------------------------------------------------------------------
' IsCOMPortAvailable (Function), source lines 260-274
' -----------------------------------------------------------------------------
Private Function IsCOMPortAvailable(ByVal PortName As String) As Boolean
    On Error GoTo 10
    PortName = Trim(PortName)
    PortName = RemoveTrailingColon(PortName)
    Dim Handle As LongPtr
    Handle = CreateFileA("\\.\" & PortName, GENERIC_READ, 0, 0, OPEN_EXISTING, 0, 0)
    
    If Handle = INVALID_HANDLE_VALUE Then
        IsCOMPortAvailable = False
    Else
        IsCOMPortAvailable = True
        CloseHandle Handle
    End If
10
End Function

' -----------------------------------------------------------------------------
' RemoveTrailingColon (Function), source lines 275-283
' -----------------------------------------------------------------------------
Private Function RemoveTrailingColon(ByVal s As String) As String
    On Error GoTo 10
    If right(s, 1) = ":" Then
        RemoveTrailingColon = left(s, Len(s) - 1)
    Exit Function
    End If
10
    RemoveTrailingColon = s
End Function

' -----------------------------------------------------------------------------
' Sort (Sub), source lines 284-301
' -----------------------------------------------------------------------------
Private Sub Sort(arr() As deviceInfo)
    If Not Not arr Then
        Dim i As Long
        Dim fSwitched As Boolean
        Dim arrTemp As deviceInfo
        Do
            fSwitched = False
            For i = 0 To UBound(arr) - 1
                If addZerroForDigits(arr(i).name) > addZerroForDigits(arr(i + 1).name) Then
                    fSwitched = True
                    arrTemp = arr(i)
                    arr(i) = arr(i + 1)
                    arr(i + 1) = arrTemp
                End If
            Next
        Loop While fSwitched
    End If
End Sub

' -----------------------------------------------------------------------------
' addZerroForDigits (Function), source lines 302-362
' -----------------------------------------------------------------------------
Private Function addZerroForDigits(ByVal txt As String, Optional ByVal maxDigits As Long = 50) As String
    txt = Replace(txt, ",", ".")
    Dim q As Long
    Dim code As Integer
    Dim lastCode As Integer
    Dim isStartingDigit As Boolean
    Dim startPositionDigits() As Long
    Dim endPositionDigits() As Long
    Dim digitsCount As Long
    
    Dim toRight() As Boolean ' (если дробь, то добавить нули справа)
    
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

' -----------------------------------------------------------------------------
' AddToArray (Sub), source lines 363-373
' -----------------------------------------------------------------------------
Private Sub AddToArray(ByRef arr() As deviceInfo, newItem As deviceInfo)
    Dim NewIndex As Long
    If Not Not arr() Then
        NewIndex = UBound(arr) + 1
        ReDim Preserve arr(NewIndex)
        arr(NewIndex) = newItem
    Else
        ReDim arr(0)
        arr(0) = newItem
    End If
End Sub

' -----------------------------------------------------------------------------
' Configure (Sub), source lines 374-395
' -----------------------------------------------------------------------------
Public Sub Configure(PortName As String, hwnd As LongPtr)
    
    If InStr(1, PortName, "COM", vbTextCompare) Then
        COMSettingsForm.PortName = RemoveTrailingColon(Trim(PortName))
        COMSettingsForm.AdvancedSettings = False
        COMSettingsForm.deviceId = ""
        COMSettingsForm.Show 1
        Exit Sub
    End If
    

    Dim retVal As Long
    retVal = ConfigurePort(vbNullString, hwnd, PortName)
    
    If retVal = 0 Then
        If Err.LastDllError = ERROR_CANCELLED Then
            'MsgBox "The dialog was canceled.", vbInformation, "Dialog closed"
        Else
            MsgBox "The port is not available", vbInformation, "Error " & Err.LastDllError
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' PtrToString (Function), source lines 397-400
' -----------------------------------------------------------------------------
Private Function PtrToString(ByVal StringPtr As LongPtr) As String
    PtrToString = Space(lstrlen(StringPtr))
    lstrcpy PtrToString, StringPtr
End Function

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 401-406
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    Dim q As Long
    For q = 0 To lCount - 1
        If Ports(q).Initialized Then DeInitialize (q)
    Next
End Sub
