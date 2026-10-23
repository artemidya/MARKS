Private Sub Initialize(i As Long)

    With Interfaces(i)
        .Init = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer
        .TERM = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator
        .Initialized = True
        
        PlotterForm.AddLog .Init
        Dim bytes() As Byte
        bytes = StrToAscii(.Init)
        
        If Not Not bytes() Then
            Dim usbHandle As Long
            Dim dataToSend As String
            Dim BytesWritten As Long
        
            If .isComPortAdapter And InStr(1, Interfaces(i).PortName, "COM", vbTextCompare) Then
                'Reset COM
                COM.CommOpen 1, .FullDevicePath, RemoveTrailingColon(Trim(.PortName))
                COM.CommClose 1
            End If
            ' Открываем USB-порт
            usbHandle = CreateFile(.FullDevicePath, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL + FILE_FLAG_SEQUENTIAL_SCAN, 0)
            
            If usbHandle <> -1 Then
                ' Отправляем данные
                Dim size As Long
                size = UBound(bytes()) + 1
                WriteFile usbHandle, bytes(0), size, BytesWritten, 0
                CloseHandle usbHandle
            Else
        '        MsgBox "Ошибка открытия USB-порта!", vbCritical
            End If
        End If
    End With
End Sub
