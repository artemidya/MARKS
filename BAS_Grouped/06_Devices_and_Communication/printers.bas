' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: printers.cls
' Procedures: 21
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 107-109
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Refresh
End Sub

' -----------------------------------------------------------------------------
' Connected (PropertyGet), source lines 110-112
' -----------------------------------------------------------------------------
Public Property Get Connected(ByVal i As Long) As Boolean
    Connected = printers(i).Active
End Property

' -----------------------------------------------------------------------------
' VidAndPid (PropertyGet), source lines 113-123
' -----------------------------------------------------------------------------
Property Get VidAndPid(OutPutName As String) As String
    Dim i As Long
    i = GetIndex(OutPutName)
    
    With printers(i)
        If .VidAndPid = "" Then
           .VidAndPid = VidAndPidByPrinterName(OutPutName)
        End If
        VidAndPid = .VidAndPid
    End With
End Property

' -----------------------------------------------------------------------------
' VidAndPidByIndex (PropertyGet), source lines 124-131
' -----------------------------------------------------------------------------
Property Get VidAndPidByIndex(ByVal index As Long) As String
    With printers(index)
        If .VidAndPid = "" Then
           .VidAndPid = VidAndPidByPrinterName(.name)
        End If
        VidAndPidByIndex = .VidAndPid
    End With
End Property

' -----------------------------------------------------------------------------
' VidAndPidByPrinterName (Function), source lines 132-159
' -----------------------------------------------------------------------------
Function VidAndPidByPrinterName(PrinterName As String) As String
    On Error GoTo 10
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity where Name like '%" & PrinterName & "%'")
    
    Dim vidPos As Long, pidPos As Long
    Dim vid As String, pid As String
    
    For Each objItem In colItems
        vidPos = 0
        pidPos = 0
        vidPos = InStr(1, objItem.deviceId, "VID_", vbTextCompare)
        pidPos = InStr(1, objItem.deviceId, "PID_", vbTextCompare)
'        Debug.Print objItem.Name & " - " & objItem.deviceID
    
        If vidPos > 0 And pidPos > 0 Then
            vid = LCase(Mid(objItem.deviceId, vidPos, 8)) ' VID_XXXX = 8 символов
            pid = LCase(Mid(objItem.deviceId, pidPos, 8)) ' PID_XXXX = 8 символов
            VidAndPidByPrinterName = vid & "&" & pid
            Exit Function
        End If
    Next
10
End Function

' -----------------------------------------------------------------------------
' count (PropertyGet), source lines 166-168
' -----------------------------------------------------------------------------
Public Property Get count() As Long
    count = lCount
End Property

' -----------------------------------------------------------------------------
' Names (PropertyGet), source lines 169-171
' -----------------------------------------------------------------------------
Public Property Get Names(i As Long) As String
    Names = printers(i).name
End Property

' -----------------------------------------------------------------------------
' x (Function), source lines 172-175
' -----------------------------------------------------------------------------
Public Function x(OutPutName As String) As Long
    Dim i As Long:    i = GetIndex(OutPutName)
    x = printers(i).x
End Function

' -----------------------------------------------------------------------------
' y (Function), source lines 176-179
' -----------------------------------------------------------------------------
Public Function y(OutPutName As String) As Long
    Dim i As Long:    i = GetIndex(OutPutName)
    y = printers(i).y
End Function

' -----------------------------------------------------------------------------
' Out (Sub), source lines 180-220
' -----------------------------------------------------------------------------
Public Sub Out(OutPutName As String, ByRef bytes() As Byte, x As Long, y As Long)
    On Error GoTo 10
    Dim i As Long:    i = GetIndex(OutPutName)
    With printers(i)
        If .Init <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer _
        Or .TERM <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator Then
            If .Initialized Then DeInitialize (i)
            Initialize (i)
        End If
        .x = x
        .y = y
        Dim q As Long
        Dim written As Long
        Dim bufferLen As Long
        Const bufSize As Long = &H4000
        
        Dim di As DOC_INFO_1
        di.pDocName = "PlotCalc" 'Имя документа в очереди печати
        di.pOutputFile = vbNullString
        di.pDatatype = "RAW"
        
        If UBound(bytes) > -1 Then
            Call StartDocPrinter(.phPrn, 1, di)
            bufferLen = UBound(bytes) + 1
            Call StartPagePrinter(.phPrn)
            For q = 1 To bufferLen \ bufSize
                Call WritePrinter(.phPrn, bytes(q * bufSize - bufSize), bufSize, written)
                'MsgBox "Print Send"
            Next
            ' Get last chunk of file if it doesn't
            If bufferLen Mod bufSize Then
                Call WritePrinter(.phPrn, bytes(q * bufSize - bufSize), bufferLen Mod bufSize, written)
            End If
            
            Call EndPagePrinter(.phPrn)
            Call EndDocPrinter(.phPrn)

        End If
    End With
10
End Sub

' -----------------------------------------------------------------------------
' Initialize (Sub), source lines 221-261
' -----------------------------------------------------------------------------
Private Sub Initialize(i As Long)
    With printers(i)
        .profileIndex = PlotterForm.TabStrip1.SelectedItem.index
        .Init = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer
        .TERM = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator
        .Initialized = True
        
        Dim di As DOC_INFO_1
        di.pDocName = "PlotCalc" 'Имя документа в очереди печати
        di.pOutputFile = vbNullString
        di.pDatatype = "RAW"
        
        Dim PD As PRINTER_DEFAULTS
        Call OpenPrinter(.name, .phPrn, PD)

        Dim q As Long
        Dim written As Long
        Dim buffer() As Byte
        Dim bufferLen As Long
        Const bufSize As Long = &H4000
        
        ' Read in 16K buffers and spool.
        PlotterForm.AddLog .Init
        buffer = StrToAscii(.Init)
        
        If Not Not buffer() Then
            Call StartDocPrinter(.phPrn, 1, di)
            Call StartPagePrinter(.phPrn)
            bufferLen = UBound(buffer) + 1
            For q = 1 To bufferLen \ bufSize
                Call WritePrinter(.phPrn, buffer(q * bufSize - bufSize), bufSize, written)
            Next
            ' Get last chunk of file if it doesn't
            If bufferLen Mod bufSize Then
                Call WritePrinter(.phPrn, buffer(q * bufSize - bufSize), bufferLen Mod bufSize, written)
            End If
            Call EndPagePrinter(.phPrn)
            Call EndDocPrinter(.phPrn)
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' DeInitialize (Sub), source lines 262-301
' -----------------------------------------------------------------------------
Private Sub DeInitialize(i As Long)
    On Error GoTo 10
    With printers(i)
        Dim q As Long
        Dim written As Long
        Dim buffer() As Byte
        Dim bufferLen As Long
        Const bufSize As Long = &H4000
        
        PlotterForm.AddLog .TERM
        buffer = StrToAscii(.TERM)
        .Initialized = False
        .Init = ""
        .TERM = ""

        Dim di As DOC_INFO_1
        di.pDocName = "PlotCalc" 'Имя документа в очереди печати
        di.pOutputFile = vbNullString
        di.pDatatype = "RAW"
        

        If UBound(buffer) > -1 Then
            Call StartDocPrinter(.phPrn, 1, di)
            Call StartPagePrinter(.phPrn)
            bufferLen = UBound(buffer) + 1
            For q = 1 To bufferLen \ bufSize
                Call WritePrinter(.phPrn, buffer(q * bufSize - bufSize), bufSize, written)
            Next
            ' Get last chunk of file if it doesn't
            If bufferLen Mod bufSize Then
                Call WritePrinter(.phPrn, buffer(q * bufSize - bufSize), bufferLen Mod bufSize, written)
            End If
            Call EndPagePrinter(.phPrn)
            Call EndDocPrinter(printers(i).phPrn)
        End If
    End With
10
    ' Shut down spooling process.
    Call ClosePrinter(printers(i).phPrn)
End Sub

' -----------------------------------------------------------------------------
' GetIndex (Function), source lines 302-310
' -----------------------------------------------------------------------------
Private Function GetIndex(OutPutName As String) As Long
    Dim q As Long
    For q = 0 To UBound(printers)
        If printers(q).name = OutPutName Then
            GetIndex = q
            Exit Function
        End If
    Next
End Function

' -----------------------------------------------------------------------------
' Refresh (Sub), source lines 311-344
' -----------------------------------------------------------------------------
Public Sub Refresh()
    Class_Terminate 'DeInitialize
    Erase printers()
    lCount = 0
    On Error GoTo 10
    Dim abEnumBuffer() As Byte, cBufferSize As Long
    Dim uPrinterInfo As PRINTER_INFO, cStructSize As Long

    Dim i As Long
    Const lngLevel As Long = 1
    Const lngPrinterMax As Long = 64
    Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, lngLevel, ByVal 0&, 0, cBufferSize, lCount)
    If cBufferSize = 0 Then Exit Sub
    
    ReDim abEnumBuffer(0 To cBufferSize - 1)
    Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, lngLevel, abEnumBuffer(0), cBufferSize, cBufferSize, lCount)
    If lCount = 0 Then Exit Sub
    
    ReDim printers(lCount - 1)
    
    Dim PrintersStatus() As ConnectionStatus
    PrintersStatus() = ListPrintersStatus
    
    cStructSize = LenB(uPrinterInfo)
    For i = 0 To lCount - 1
        MoveMemory uPrinterInfo, abEnumBuffer(cStructSize * i), cStructSize
        With printers(i)
            .name = PtrToString(uPrinterInfo.pName)
            .Active = StatusFromListByName(.name, PrintersStatus())
        End With
    Next
    Sort printers()
10
End Sub

' -----------------------------------------------------------------------------
' Sort (Sub), source lines 345-362
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
' addZerroForDigits (Function), source lines 363-423
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

' -----------------------------------------------------------------------------
' StatusFromListByName (Function), source lines 424-434
' -----------------------------------------------------------------------------
Private Function StatusFromListByName(ByRef name As String, ByRef ConnectionStatuses() As ConnectionStatus) As Boolean
    Dim q As Long
    If Not Not ConnectionStatuses() Then
        For q = 0 To UBound(ConnectionStatuses())
            If name = ConnectionStatuses(q).name Then
                StatusFromListByName = ConnectionStatuses(q).Connected
                Exit Function
            End If
        Next
    End If
End Function

' -----------------------------------------------------------------------------
' ListPrintersStatus (Function), source lines 435-462
' -----------------------------------------------------------------------------
Private Function ListPrintersStatus() As ConnectionStatus()
    On Error GoTo 10
    
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Printer")

    If colItems.count > 0 Then
        Dim ConnectionStatuses() As ConnectionStatus
        ReDim ConnectionStatuses(colItems.count - 1)
    
        Dim i As Long
        
        For Each objItem In colItems
            With ConnectionStatuses(i)
    '            Debug.Print objItem.Name & " | Port: " & objItem.portName & " | Status: " & objItem.WorkOffline
                .name = objItem.name
                .Connected = Not objItem.WorkOffline
                i = i + 1
            End With
        Next
        ListPrintersStatus = ConnectionStatuses
    End If
10
End Function

' -----------------------------------------------------------------------------
' PtrToString (Function), source lines 464-467
' -----------------------------------------------------------------------------
Private Function PtrToString(ByVal StringPtr As LongPtr) As String
    PtrToString = Space(lstrlen(StringPtr))
    lstrcpy PtrToString, StringPtr
End Function

' -----------------------------------------------------------------------------
' Configure (Sub), source lines 469-485
' -----------------------------------------------------------------------------
Public Sub Configure(PrinterName As String, hwnd As LongPtr)
    On Error GoTo TERM:
    Dim iRet As Long
    Dim hPrinter As LongPtr
    Dim PD As PRINTER_DEFAULTS
    
    iRet = OpenPrinter(PrinterName, hPrinter, PD)
    If (iRet = 0) Or (hPrinter = 0) Then
       'Can't access current printer. Bail out doing nothing
       Exit Sub
    End If
    
    PrinterProperties hwnd, hPrinter
'    MsgBox pd.pDevmode.dmPaperWidth
TERM:
    If hPrinter <> 0 Then ClosePrinter hPrinter
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 487-492
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    Dim q As Long
    For q = 0 To lCount - 1
        If printers(q).Initialized Then DeInitialize (q)
    Next
End Sub
