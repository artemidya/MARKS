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
