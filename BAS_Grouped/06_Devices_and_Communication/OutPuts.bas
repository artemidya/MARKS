' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: OutPuts.cls
' Procedures: 46
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 51-63
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    'Init BlocksSokets
    BlocksSocketShared.HookForm GetForegroundWindow
    Set PlotterClient = New BlocksSocket
    PlotterClient.OnReciveOld_TurnON = True
    PlotterClientLastTabStrip1IndexValue = -1

    FilerRefresh

    Set printers = New printers
    Set Ports = New Ports
    Set Interfaces = New Interfaces
End Sub

' -----------------------------------------------------------------------------
' PlotterClient_OnConnect (Sub), source lines 64-72
' -----------------------------------------------------------------------------
Private Sub PlotterClient_OnConnect(client As BlocksSocket)
    PlotterClientConnected = True
    PlotterClientWaitingStatus = False
    PlotterClientInitialized = False
    PlotterClientInit = ""
    PlotterClientTERM = ""
    'R E
    RaiseEvent WiFiChange
End Sub

' -----------------------------------------------------------------------------
' PlotterClient_OnDisconnect (Sub), source lines 74-81
' -----------------------------------------------------------------------------
Private Sub PlotterClient_OnDisconnect(client As BlocksSocket)
    PlotterClientConnected = False
    PlotterClientWaitingStatus = False
    PlotterClientlastOperationTime = Timer
    PlotterClientInitialized = False
    'R E
    RaiseEvent WiFiChange
End Sub

' -----------------------------------------------------------------------------
' WiFi_change_check_loop (Sub), source lines 82-103
' -----------------------------------------------------------------------------
Public Sub WiFi_change_check_loop()
    'Check wifi
    If PlotterClientLastTabStrip1IndexValue <> PlotterForm.TabStrip1.SelectedItem.index Then
        PlotterClient.Disconnect
        PlotterClientlastOperationTime = Timer
        PlotterClientLastTabStrip1IndexValue = PlotterForm.TabStrip1.SelectedItem.index
        Exit Sub
    End If

    If PlotterClientWaitingStatus And (Timer - PlotterClientlastOperationTime) > 6 Then
        PlotterClient.Disconnect
        PlotterClientlastOperationTime = Timer
        Exit Sub
    End If

    '
    If PlotterClientWaitingStatus = False And PlotterClientConnected = False Then
        PlotterClientWaitingStatus = True
        PlotterClientlastOperationTime = Timer
        PlotterClient.Connect PlotterForm.RegData1.IP(PlotterForm.TabStrip1.SelectedItem.index + 1), PlotterForm.RegData1.Port(PlotterForm.TabStrip1.SelectedItem.index + 1)
    End If
End Sub

' -----------------------------------------------------------------------------
' WiFi_Refresh (Sub), source lines 104-106
' -----------------------------------------------------------------------------
Public Sub WiFi_Refresh()
    PlotterClientLastTabStrip1IndexValue = -1
End Sub

' -----------------------------------------------------------------------------
' OutTypeByIndex (PropertyGet), source lines 107-119
' -----------------------------------------------------------------------------
Property Get OutTypeByIndex(ByVal index As Long) As OutType
    If index < printers.count Then
        OutTypeByIndex = Printer
    ElseIf index < (printers.count + Interfaces.count) Then
        OutTypeByIndex = Interface
    ElseIf index < (printers.count + Interfaces.count + Ports.count) Then
        OutTypeByIndex = Port
    ElseIf index < (printers.count + Interfaces.count + Ports.count + 1) Then
        OutTypeByIndex = WiFi
    ElseIf index < (printers.count + Interfaces.count + Ports.count + 2) Then
        OutTypeByIndex = file
    End If
End Property

' -----------------------------------------------------------------------------
' ExparedID (PropertyGet), source lines 120-124
' -----------------------------------------------------------------------------
Property Get ExparedID(OutPutName As String) As String
    If OutType(OutPutName) = Interface Then
            ExparedID = Interfaces.ExparedID(OutPutName)
    End If
End Property

' -----------------------------------------------------------------------------
' ExparedIDByIndex (PropertyGet), source lines 126-131
' -----------------------------------------------------------------------------
Property Get ExparedIDByIndex(ByVal index As Long) As String
    If (index >= printers.count) And (index < (printers.count + Interfaces.count)) Then
        index = index - printers.count
        ExparedIDByIndex = Interfaces.ExparedIDByIndex(index)
    End If
End Property

' -----------------------------------------------------------------------------
' VidAndPid (PropertyGet), source lines 133-144
' -----------------------------------------------------------------------------
Property Get VidAndPid(OutPutName As String) As String
    Dim mOutType As OutType
    mOutType = OutType(OutPutName)
    Select Case mOutType
        Case Printer
            VidAndPid = printers.VidAndPid(OutPutName)
        Case Interface
            VidAndPid = Interfaces.VidAndPid(OutPutName)
        Case Port
            VidAndPid = GetVidAndPidFromVirtualPorts(OutPutName)
    End Select
End Property

' -----------------------------------------------------------------------------
' VidAndPidByIndex (PropertyGet), source lines 145-157
' -----------------------------------------------------------------------------
Property Get VidAndPidByIndex(ByVal index As Long) As String
    If index < printers.count Then
        VidAndPidByIndex = printers.VidAndPidByIndex(index)
    ElseIf index < (printers.count + Interfaces.count) Then
        index = index - printers.count
        VidAndPidByIndex = Interfaces.VidAndPidByIndex(index)
    ElseIf index < (printers.count + Interfaces.count + Ports.count) Then
        'Ports
        index = index - printers.count - Interfaces.count
        VidAndPidByIndex = GetVidAndPidFromVirtualPorts(Ports.Names(index))
    Else
    End If
End Property

' -----------------------------------------------------------------------------
' GetVidAndPidFromVirtualPorts (Function), source lines 158-169
' -----------------------------------------------------------------------------
Private Function GetVidAndPidFromVirtualPorts(ByVal Port As String) As String
    Port = UCase(RemoveTrailingColon(Port))
    Dim q As Long
    For q = 0 To Interfaces.count - 1
        If Interfaces.isComOrLPT_Adapter(q) Then
            If UCase(Interfaces.PortName(q)) = Port Then
                GetVidAndPidFromVirtualPorts = Interfaces.VidAndPidByIndex(q)
                Exit Function
            End If
        End If
    Next
End Function

' -----------------------------------------------------------------------------
' RemoveTrailingColon (Function), source lines 170-176
' -----------------------------------------------------------------------------
Function RemoveTrailingColon(ByVal s As String) As String
    If right(s, 1) = ":" Then
        RemoveTrailingColon = left(s, Len(s) - 1)
    Else
        RemoveTrailingColon = s
    End If
End Function

' -----------------------------------------------------------------------------
' Ymm (PropertyGet), source lines 177-201
' -----------------------------------------------------------------------------
Public Property Get Ymm() As Double
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Dim y As Double
    Select Case mOutType
        Case Printer
            y = printers.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Interface
            y = Interfaces.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Port
            y = Ports.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case WiFi
            y = WiFiCoord.y
        Case file
            y = FileCoord.y
    End Select
    
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        y = y / Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale) / 40
    Else
        y = y / Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale) / 40
    End If
    
    Ymm = y
End Property

' -----------------------------------------------------------------------------
' Xmm (PropertyGet), source lines 202-226
' -----------------------------------------------------------------------------
Public Property Get Xmm() As Double
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Dim x As Double
    Select Case mOutType
        Case Printer
            x = printers.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Interface
            x = Interfaces.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Port
            x = Ports.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case WiFi
            x = WiFiCoord.x
        Case file
            x = FileCoord.x
    End Select
    
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        x = x / Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale) / 40
    Else
        x = x / Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale) / 40
    End If
    
    Xmm = x
End Property

' -----------------------------------------------------------------------------
' y (PropertyGet), source lines 227-242
' -----------------------------------------------------------------------------
Public Property Get y() As Long
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Select Case mOutType
        Case Printer
            y = printers.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Interface
            y = Interfaces.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Port
            y = Ports.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case WiFi
            y = WiFiCoord.y
        Case file
            y = FileCoord.y
    End Select
End Property

' -----------------------------------------------------------------------------
' x (PropertyGet), source lines 243-258
' -----------------------------------------------------------------------------
Public Property Get x() As Long
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Select Case mOutType
        Case Printer
            x = printers.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Interface
            x = Interfaces.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Port
            x = Ports.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case WiFi
            x = WiFiCoord.x
        Case file
            x = FileCoord.x
    End Select
End Property

' -----------------------------------------------------------------------------
' MoveDown (Sub), source lines 261-291
' -----------------------------------------------------------------------------
Public Sub MoveDown(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dx = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
    Else
        dy = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
    End If
    
    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If
    
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub

' -----------------------------------------------------------------------------
' MoveUP (Sub), source lines 292-322
' -----------------------------------------------------------------------------
Public Sub MoveUP(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dx = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
    Else
        dy = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
    End If
    
    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If
    
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub

' -----------------------------------------------------------------------------
' MoveRight (Sub), source lines 323-354
' -----------------------------------------------------------------------------
Public Sub MoveRight(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dy = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    Else
        dx = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    End If
    
    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If
    
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub

' -----------------------------------------------------------------------------
' MoveLeft (Sub), source lines 355-385
' -----------------------------------------------------------------------------
Public Sub MoveLeft(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dy = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    Else
        dx = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    End If

    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If

    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub

' -----------------------------------------------------------------------------
' MoveLeftUp (Sub), source lines 386-418
' -----------------------------------------------------------------------------
Public Sub MoveLeftUp(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dy = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
        dx = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
    Else
        dx = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
        dy = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
    End If

    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If

    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y, 30)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub

' -----------------------------------------------------------------------------
' MoveUpRight (Sub), source lines 419-451
' -----------------------------------------------------------------------------
Public Sub MoveUpRight(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dx = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
        dy = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    Else
        dy = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
        dx = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    End If
    
    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If
    
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y, 30)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub

' -----------------------------------------------------------------------------
' MoveRightDown (Sub), source lines 452-485
' -----------------------------------------------------------------------------
Public Sub MoveRightDown(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dy = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
        dx = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
    Else
        dx = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
        dy = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
    End If
    
    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If
    
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y, 30)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub

' -----------------------------------------------------------------------------
' MoveDownLeft (Sub), source lines 486-518
' -----------------------------------------------------------------------------
Public Sub MoveDownLeft(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dx = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
        dy = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    Else
        dy = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
        dx = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    End If
    
    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If
    
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y, 30)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub

' -----------------------------------------------------------------------------
' Move00 (Sub), source lines 519-530
' -----------------------------------------------------------------------------
Public Sub Move00()
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(0, 0)
            Out StrToAscii(txt), 0, 0
        End If
    Else
        FileCoord.x = 0
        FileCoord.y = 0
    End If
End Sub

' -----------------------------------------------------------------------------
' test (Sub), source lines 531-537
' -----------------------------------------------------------------------------
Public Sub test()
    Move00
    MoveLeft 10
    MoveUP 10
    MoveRight 10
    MoveDown 10
End Sub

' -----------------------------------------------------------------------------
' ComSettingsStr (PropertyGet), source lines 538-549
' -----------------------------------------------------------------------------
Property Get ComSettingsStr() As String
    On Error GoTo 10
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Select Case mOutType
        Case Interface
            ComSettingsStr = Interfaces.ComSettingsStr(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Port
            ComSettingsStr = Ports.ComSettingsStr(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    End Select
10
End Property

' -----------------------------------------------------------------------------
' Out (Function), source lines 551-586
' -----------------------------------------------------------------------------
Public Function Out(ByRef bytes() As Byte, x As Long, y As Long, Optional dntInitialize As Boolean = False) As Boolean
    On Error GoTo 10
    Dim mOutType As OutType
    Dim f As Boolean
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Select Case mOutType
        Case Printer
            printers.Out KillBirdsAndSpaces(PlotterForm.Combo1.text), bytes, x, y
            Out = True
        Case Interface
            Interfaces.Out KillBirdsAndSpaces(PlotterForm.Combo1.text), bytes, x, y, dntInitialize
            Out = True
        Case Port
            Ports.Out KillBirdsAndSpaces(PlotterForm.Combo1.text), bytes, x, y, dntInitialize
            Out = True
        Case WiFi
            WiFiCoord.x = x
            WiFiCoord.y = y
            OutToWiFi bytes()
            Out = True
        Case file
            FileCoord.x = x
            FileCoord.y = y
            f = OutToFile(bytes)
            Out = f
    End Select
    
    If (mOutType = file And f) Or mOutType <> file Then
        If UBound(bytes) < 7000 Then
            PlotterForm.AddLog PlotterLanguage.AsciiToString(bytes)
        Else
            PlotterForm.AddLog "Write " & UBound(bytes) + 1 & " bytes"
        End If
    End If
10
End Function

' -----------------------------------------------------------------------------
' OutToWiFi (Sub), source lines 587-595
' -----------------------------------------------------------------------------
Private Sub OutToWiFi(ByRef bytes() As Byte)
    If (PlotterClientInit <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer _
    Or PlotterClientTERM <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator) Then
        If PlotterClientInitialized Then DeInitialize
        Initialize
    End If
        
    PlotterClient.SendOld bytes()
End Sub

' -----------------------------------------------------------------------------
' Initialize (Sub), source lines 596-611
' -----------------------------------------------------------------------------
Private Sub Initialize()
    On Error GoTo 10
    PlotterClientInit = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer
    PlotterClientTERM = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator
    PlotterClientInitialized = True
    
    PlotterForm.AddLog PlotterClientInit
    Dim bytes() As Byte
    bytes = StrToAscii(PlotterClientInit)
    
    If UBound(bytes) > -1 Then
        PlotterClient.SendOld bytes()
    End If
    
10
End Sub

' -----------------------------------------------------------------------------
' DeInitialize (Sub), source lines 612-628
' -----------------------------------------------------------------------------
Private Sub DeInitialize()

    On Error GoTo 10
    PlotterForm.AddLog PlotterClientTERM
    Dim bytes() As Byte
    bytes = StrToAscii(PlotterClientTERM)

    PlotterClientInitialized = False
    PlotterClientInit = ""
    PlotterClientTERM = ""

    
    If UBound(bytes) > -1 Then
        PlotterClient.SendOld bytes()
    End If
10
End Sub

' -----------------------------------------------------------------------------
' count (PropertyGet), source lines 630-632
' -----------------------------------------------------------------------------
Public Property Get count() As Long
    count = printers.count + Interfaces.count + Ports.count + 2
End Property

' -----------------------------------------------------------------------------
' Names (PropertyGet), source lines 633-647
' -----------------------------------------------------------------------------
Public Property Get Names(ByVal i As Long) As String
    If i < printers.count Then
        Names = printers.Names(i)
    ElseIf i < (printers.count + Interfaces.count) Then
        i = i - printers.count
        Names = Interfaces.Names(i)
    ElseIf i < (printers.count + Interfaces.count + Ports.count) Then
        i = i - printers.count - Interfaces.count
        Names = Ports.Names(i)
    ElseIf i = (printers.count + Interfaces.count + Ports.count) Then
        Names = "Wi-Fi / Ethernet"
    Else
        Names = "File (flash card)"
    End If
End Property

' -----------------------------------------------------------------------------
' Connected (PropertyGet), source lines 648-664
' -----------------------------------------------------------------------------
Public Property Get Connected(ByVal i As Long) As Boolean
    If i < printers.count Then
        Connected = printers.Connected(i)
    ElseIf i < (printers.count + Interfaces.count) Then
        i = i - printers.count
        Connected = Interfaces.Connected(i)
    ElseIf i < (printers.count + Interfaces.count + Ports.count) Then
        i = i - printers.count - Interfaces.count
        Connected = Ports.Connected(i)
    ElseIf i = (printers.count + Interfaces.count + Ports.count) Then
        
        Connected = PlotterClientConnected '
    Else
        Connected = FileConnectedIfUSB ' "File (flash card)"
    End If

End Property

' -----------------------------------------------------------------------------
' OutTypeToString (Function), source lines 665-680
' -----------------------------------------------------------------------------
Public Function OutTypeToString(OutPutName As String) As String
    Dim mOutType As OutType
    mOutType = OutType(OutPutName)
    Select Case mOutType
        Case Printer
            OutTypeToString = "Printer"
        Case Interface
            OutTypeToString = "Interface"
        Case Port
            OutTypeToString = "Port"
        Case WiFi
            OutTypeToString = "WiFi"
        Case file
            OutTypeToString = "File"
    End Select
End Function

' -----------------------------------------------------------------------------
' OutToFile (Function), source lines 683-730
' -----------------------------------------------------------------------------
Private Function OutToFile(ByRef bytes() As Byte) As Boolean

    Dim fileExtension As String
    fileExtension = LoadSaveDialog.GetFileExtension(Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT)
    If fileExtension = "" Then fileExtension = "*"
 
    Dim filter As LoadSaveDialog.filter
    With filter
        If UCase(fileExtension) = "PLT" Then
            ReDim .FilterRows(1)
            .FilterRows(0).Mask = "*.PLT"
            .FilterRows(0).Description = "Plotter file (*.plt)"
            .FilterRows(1).Mask = "*.*"
            .FilterRows(1).Description = "All files (*.*)"
        Else
            ReDim .FilterRows(2)
            .FilterRows(0).Mask = "*." & UCase(fileExtension)
            .FilterRows(0).Description = PlotterForm.Text4.text & " (*." & LCase(fileExtension) & ")"
            .FilterRows(1).Mask = "*.PLT"
            .FilterRows(1).Description = "Plotter file (*.plt)"
            .FilterRows(2).Mask = "*.*"
            .FilterRows(2).Description = "All files (*.*)"
        End If
        Dim strFilter As String
        Dim q As Long
        For q = 0 To UBound(.FilterRows)
            strFilter = strFilter & .FilterRows(q).Description & "|" & .FilterRows(q).Mask & "|"
        Next
    End With
    
    Dim result As String
'    result = LoadSaveDialog.SaveFile(Filter, PlotterForm.RegData1.LastPlotterDirectory, Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT)
    On Error GoTo 20
    result = Application.CorelScriptTools.GetFileBox(strFilter, , 1, Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT, , PlotterForm.RegData1.LastPlotterDirectory)
20
    If result <> "" Then
        Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT = LoadSaveDialog.GetFile(result)
        PlotterForm.RegData1.LastPlotterDirectory = LoadSaveDialog.GetPath(result)

        If fFileLen(result) > UBound(bytes) + 1 Then KillFile result
        Dim fileNumber As Long
        fileNumber = FreeFile
        Open result For Binary Access Write As fileNumber
            Put fileNumber, , bytes()
        Close fileNumber
        OutToFile = True
    End If
End Function

' -----------------------------------------------------------------------------
' Refresh (Sub), source lines 731-737
' -----------------------------------------------------------------------------
Public Sub Refresh()
    printers.Refresh
    Interfaces.Refresh
    Ports.Refresh
    'file
    FilerRefresh
End Sub

' -----------------------------------------------------------------------------
' FilerRefresh (Sub), source lines 738-744
' -----------------------------------------------------------------------------
Private Sub FilerRefresh()
    If IsDriveFixed(PlotterForm.RegData1.LastPlotterDirectory) Then
        FileConnectedIfUSB = True
    Else
        FileConnectedIfUSB = FolderExists(PlotterForm.RegData1.LastPlotterDirectory)
    End If
End Sub

' -----------------------------------------------------------------------------
' FolderExists (Function), source lines 745-753
' -----------------------------------------------------------------------------
Private Function FolderExists(ByVal FolderPath As String) As Boolean
    On Error GoTo 10
    If right(FolderPath, 1) = "\" Then
        FolderPath = left(FolderPath, Len(FolderPath) - 1)
    End If
    
    FolderExists = (Dir(FolderPath, vbDirectory) <> "")
10
End Function

' -----------------------------------------------------------------------------
' IsDriveFixed (Function), source lines 755-768
' -----------------------------------------------------------------------------
Private Function IsDriveFixed(ByVal FilePath As String) As Boolean
    If FilePath = "" Then Exit Function
    On Error GoTo 10
    
    IsDriveFixed = True
    Dim DriveLetter As String
    DriveLetter = UCase(left(FilePath, 3)) ' например: "C:\"
    
    Dim driveType As Long
    driveType = GetDriveType(DriveLetter)
    
    IsDriveFixed = (driveType = DRIVE_FIXED)
10
End Function

' -----------------------------------------------------------------------------
' Configure (Sub), source lines 769-789
' -----------------------------------------------------------------------------
Public Sub Configure(OutPutName As String, hwnd As LongPtr)
    Dim mOutType As OutType
    mOutType = OutType(OutPutName)
    
    Select Case mOutType
        Case Printer
            printers.Configure OutPutName, hwnd
        Case Interface
            Interfaces.Configure OutPutName, hwnd
        Case Port
            Ports.Configure OutPutName, hwnd
        Case WiFi
            WiFiSettingsForm.Show 1
        Case file
            Dim newFile As String
            newFile = InputBox(localization.ActiveLng.Item("OutPuts.InputBox.Enter"), "PlotCalc", Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT)
            If newFile <> "" Then
                Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT = newFile
            End If
    End Select
End Sub

' -----------------------------------------------------------------------------
' OutType (Function), source lines 790-821
' -----------------------------------------------------------------------------
Private Function OutType(OutPutName As String) As OutType
    Dim q As Long
    For q = 0 To printers.count - 1
        If printers.Names(q) = OutPutName Then
            OutType = Printer
            Exit Function
        End If
    Next
    
    For q = 0 To Interfaces.count - 1
        If Interfaces.Names(q) = OutPutName Then
            OutType = Interface
            Exit Function
        End If
    Next
    
    For q = 0 To Ports.count - 1
        If Ports.Names(q) = OutPutName Then
            OutType = Port
            Exit Function
        End If
    Next
    
    If OutPutName = "Wi-Fi / Ethernet" Then
        OutType = WiFi
        Exit Function
    End If
        
    '-----
    'FILE
    OutType = file
End Function

' -----------------------------------------------------------------------------
' KillFile (Sub), source lines 822-826
' -----------------------------------------------------------------------------
Private Sub KillFile(fileName As String)
    On Error GoTo 10
    Kill fileName
10
End Sub

' -----------------------------------------------------------------------------
' fFileLen (Function), source lines 827-832
' -----------------------------------------------------------------------------
Private Function fFileLen(fileName As String) As Long
    On Error GoTo 10
    fFileLen = FileLen(fileName)
    Exit Function
10
End Function

' -----------------------------------------------------------------------------
' KillBirdsAndSpaces (Function), source lines 833-852
' -----------------------------------------------------------------------------
Private Function KillBirdsAndSpaces(ByVal s As String) As String
    ' ”далить 4 пробела в начале, если строка достаточно длинна€
    If Len(s) >= 4 Then
        If left(s, 4) = "    " Then
            KillBirdsAndSpaces = Mid(s, 5)
            Exit Function
        End If
    End If

    ' ”далить символ ChrW(10003) (?) в начале, если строка не пуста€
    If Len(s) >= 1 Then
        If left(s, 1) = ChrW(10003) Then
            KillBirdsAndSpaces = Mid(s, 2)
            Exit Function
        End If
    End If

    ' ¬ернуть как есть, если ничего не найдено
    KillBirdsAndSpaces = s
End Function

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 855-858
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    If PlotterClientInitialized Then DeInitialize
    PlotterClient.Disconnect
End Sub
