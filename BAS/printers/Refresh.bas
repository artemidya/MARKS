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
