Private Sub DrawRazm(ByVal matW As Double, ByVal matH As Double, _
                 ByVal plotW As Double, ByVal plotH As Double, _
                 ByVal uF As Double, ByVal lF As Double, _
                 ByVal rF As Double, ByVal dF As Double)
    On Error Resume Next
    If CountResults > 0 Then Command3.Enabled = True
    Command6.Enabled = False
    SelectedResult = CountResults + 1
    Command4.Enabled = False
    
    If Check2.value Then Cls True: Exit Sub
    
    'Рисует разметку только для ознакомления
    Dim wpix As Double
    Dim hpix As Double
    Dim ks As Double
    Dim dx As Double
    Dim dy As Double
    Dim oUF As Double, oLF As Double
    Dim oRF As Double, oDF As Double
    
    If matW = 0 Or matH = 0 Then Cls True: Exit Sub
    
    FillRect PaperBrush, 0, 0, WindowWidth, WindowHeigth
    DrawRect mPen, 0, 0, WindowWidth, WindowHeigth
    
    If matW / WindowWidth > matH / WindowHeigth Then
        ks = (WindowWidth - 15) / matW
    Else
        ks = (WindowHeigth - 20) / matH
    End If
    
    dx = 6
    dy = 8
    wpix = ks * plotW
    hpix = ks * plotH
    oUF = ks * uF
    oLF = ks * lF
    oRF = ks * rF
    oDF = ks * dF
    
    If wpix > ks * matW Then wpix = ks * matW
    If hpix > ks * matH Then hpix = ks * matH
    Dim wpix2 As Double
    wpix2 = wpix
    If Check3.value Then
        DrawRect PagesPen, dx, WindowHeigth - dy - hpix, wpix, hpix
    End If
    If Check5.value Then
        'Fields
        DrawRect FieldsPen, dx + oLF, WindowHeigth - dy - hpix + oUF, wpix - oRF - oLF, hpix - oUF - oDF
    End If
    
    wpix = ks * matW
    hpix = ks * matH
    If Check4.value Then
        DrawRect MaterialPen, dx, WindowHeigth - dy - hpix, wpix, hpix
    End If
    
    If Check6.value Then
        'Text
        Dim str As String
        str = Round(wpix2 / ks, 0)
        TextOut HDCForm, shiftX + wpix + dx + 2, shiftY + WindowHeigth - dy - 10, str, Len(str)
    End If
End Sub
