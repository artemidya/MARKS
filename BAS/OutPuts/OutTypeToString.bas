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
