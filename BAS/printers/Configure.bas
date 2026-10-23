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
