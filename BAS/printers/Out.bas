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
