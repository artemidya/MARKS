Private Sub Class_Terminate()
     If SizeIsSet Then
        SelectObject lhDC, oldhDC
        DeleteObject hBM
        DeleteDC lhDC
     End If
End Sub
