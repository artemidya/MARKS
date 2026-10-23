Private Sub FirstRunSub()
    'Первый запуск
    RaskroyForm.Repaint
'    UngroupOneElementGroupsStart 'Delete 1-group elements for PowerClip
    TakeData
    CurveData = PrepareToSend(RCurves()) 'Подготавливает данные к отправке, конвертирует всё в Bytes
    
    If Break = True Then Exit Sub
    
    Label1.Caption = ""
    ProgressBar1.value = 0
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
End Sub
