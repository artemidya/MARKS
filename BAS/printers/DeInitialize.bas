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
