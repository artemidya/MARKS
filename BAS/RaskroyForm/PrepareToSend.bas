Private Function PrepareToSend(ByRef RCurves() As RCurve) As Byte()
    On Error GoTo Err
    Dim q1 As Long, q2 As Long, q3 As Long
    Dim kof As Long
    '—ортируем данные
    Dim q As Long
    Dim Keys() As Long
    Dim wi() As Double
    Dim he() As Double
    ReDim Keys(1 To UBound(RCurves))
    ReDim wi(1 To UBound(RCurves))
    ReDim he(1 To UBound(RCurves))
    For q = 1 To UBound(RCurves)
        Keys(q) = q
        wi(q) = RCurves(q).Width
        he(q) = RCurves(q).Height
    Next
    
    Dim i As Long
    Dim fSwitched As Boolean
    Dim strTempW As Double
    Dim strTempH As Double
    Dim keyTemp As Long
    Do
        fSwitched = False
        For i = 1 To UBound(RCurves) - 1
            If wi(i) * he(i) < wi(i + 1) * he(i + 1) Then 'сортируем по возрастанию
                fSwitched = True
                strTempW = wi(i)
                strTempH = he(i)
                keyTemp = Keys(i)
                
                wi(i) = wi(i + 1)
                he(i) = he(i + 1)
                Keys(i) = Keys(i + 1)
                
                wi(i + 1) = strTempW
                he(i + 1) = strTempH
                Keys(i + 1) = keyTemp
            End If
        Next
    Loop While fSwitched

    Dim keysBuff As New buffer
    keysBuff.InterPlatformNetConverter = True
    keysBuff.size = UBound(Keys) * 4 + 4
    
    keysBuff.bPut_Long UBound(Keys)
    For q = 1 To UBound(Keys)
        keysBuff.bPut_Long Keys(q)
    Next

    
    AlfaWidth = RCurves(Keys(1)).Width
    AlfaHeight = RCurves(Keys(1)).Height
    NumberOfMaxFigure = Keys(1)
    
    '==============================================
    Dim curveBuffers() As buffer
    ReDim curveBuffers(1 To UBound(Keys))
    
    Dim qq2 As Long
    Dim KeysQ1 As Long
    
    For q1 = 1 To UBound(RCurves)
        ProgressBar1.value = Int(q1 / UBound(RCurves) * 10000)
        If Break = True Then Exit Function
        DoEvents
        KeysQ1 = Keys(q1)
        Set curveBuffers(KeysQ1) = New buffer
        curveBuffers(KeysQ1).InterPlatformNetConverter = True
        
        With RCurves(KeysQ1)
            If q1 > 1 Then
                qq2 = Keys(q1 - 1)
            Else
                GoTo e
            End If

            If Odinacovo(RCurves(KeysQ1), RCurves(qq2)) Then
                curveBuffers(KeysQ1).size = 4
                curveBuffers(KeysQ1).bPut_Long 1
            Else
e:
                'ј иначе если фигура нетака€ же как и прошла€ то добавить
                curveBuffers(KeysQ1).size = CalculateSize(.RPaths)
                curveBuffers(KeysQ1).bPut_Long 0
                kof = Int(.MaxLenght * 100000000000#)
                curveBuffers(KeysQ1).bPut_Long kof
                curveBuffers(KeysQ1).bPut_Long UBound(.RPaths)
                For q2 = 1 To UBound(.RPaths)
                    curveBuffers(KeysQ1).bPut_Long UBound(.RPaths(q2).RSegments)
                    For q3 = 1 To UBound(.RPaths(q2).RSegments)
                        With .RPaths(q2).RSegments(q3)
                            curveBuffers(KeysQ1).bPut_Long CLng(.SNPx)
                            curveBuffers(KeysQ1).bPut_Long CLng(.SNPy)
                            curveBuffers(KeysQ1).bPut_Long CLng(.SCPx)
                            curveBuffers(KeysQ1).bPut_Long CLng(.SCPy)
                            curveBuffers(KeysQ1).bPut_Long CLng(.ECPx)
                            curveBuffers(KeysQ1).bPut_Long CLng(.ECPy)
                        End With
                    Next
                Next
            End If
        End With
    Next
    'OKON =============================
    Dim fullBufferSize As Long
    fullBufferSize = keysBuff.size
    For q1 = 1 To UBound(RCurves)
        fullBufferSize = fullBufferSize + curveBuffers(q1).size
    Next
    
    Dim fullBuffer As New buffer
    fullBuffer.size = fullBufferSize + 4
    fullBuffer.InterPlatformNetConverter = True
    fullBuffer.bPut_Long Commands.CurveData
    fullBuffer.bPut_Bytes keysBuff.bytes
    For q1 = 1 To UBound(RCurves)
        fullBuffer.bPut_Bytes curveBuffers(q1).bytes
    Next
    PrepareToSend = fullBuffer.bytes
Err:
End Function
