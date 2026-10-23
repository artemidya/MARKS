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
