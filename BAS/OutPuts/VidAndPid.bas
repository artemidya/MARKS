Property Get VidAndPid(OutPutName As String) As String
    Dim mOutType As OutType
    mOutType = OutType(OutPutName)
    Select Case mOutType
        Case Printer
            VidAndPid = printers.VidAndPid(OutPutName)
        Case Interface
            VidAndPid = Interfaces.VidAndPid(OutPutName)
        Case Port
            VidAndPid = GetVidAndPidFromVirtualPorts(OutPutName)
    End Select
End Property
