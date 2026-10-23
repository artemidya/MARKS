Private Function DeviceIsPresentInMySystem(pi As PlotterInfo) As Boolean
    If pi.VidAndPid <> "" Then
        Dim q As Long
        If Not Not VidAndPid() Then
            With pi
                For q = 0 To UBound(VidAndPid)
                    If VidAndPid(q) = UCase(.VidAndPid) Then
                        DeviceIsPresentInMySystem = True
                        Exit For
                    End If
                Next
            End With
        End If
    End If
End Function
