Public Sub CreateLabirintShape()
    UserData.Load
    On Error GoTo 10
    
    If Application.ActiveSelection.shapes.count = 0 Then Exit Sub

    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrInch
    
    Dim Curves() As curve
    
    Application.Optimization = True
    Application.ActiveSelection.ConvertToCurves
    Application.ActiveSelection.UngroupAllEx
    GetCurves Curves()          'Преобразовуем кривые для обработки
    ActiveDocument.Undo 2
    IsInner Curves()            'Определяет является ли подпуть внутренним
    Tops Curves()               'Определяет точки-соединения
    Find Curves()               'Локальное соединение
    GlobConnectH Curves()       'Глобальное горизонтальное соединение
    GlobConnectW Curves()       'Глобальное вертикальное соединение
    OverENDconnect_H Curves()   'Последнее горизонтальное соединение
    OverENDconnect_W Curves()   'Последнее вертикальное соединение
    DrawCurves Curves()         'нарисовать кривые
    
10
    Application.ActiveDocument.Unit = lastUnit
    Application.Optimization = False
    Application.Refresh
    

End Sub
