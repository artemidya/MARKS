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
