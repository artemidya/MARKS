Private Function PressedKey() As PressedKeys

    Static myKey As Integer
    Dim mL As Boolean
    Dim mR As Boolean
    Dim mU As Boolean
    Dim mD As Boolean
    Static mLeft As Boolean
    Static mRight As Boolean
    Static mUp As Boolean
    Static mDown As Boolean
    
    
    myKey = GetAsyncKeystate(vbKeyUp)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mU = True
    
    myKey = GetAsyncKeystate(vbKeyLeft)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mL = True

    myKey = GetAsyncKeystate(vbKeyRight)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mR = True

    myKey = GetAsyncKeystate(vbKeyDown)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mD = True

    myKey = GetAsyncKeystate(vbKeyNumpad4)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mL = True

    myKey = GetAsyncKeystate(vbKeyNumpad6)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mR = True

    myKey = GetAsyncKeystate(vbKeyNumpad8)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mU = True

    myKey = GetAsyncKeystate(vbKeyNumpad2)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mD = True

    mLeft = mL And (Not mR)
    mRight = mR And (Not mL)
    mUp = mU And (Not mD)
    mDown = mD And (Not mU)
    
    
    If mLeft And mUp Then PressedKey = KeyLeftUp: Exit Function
    If mUp And mRight Then PressedKey = KeyUpRight: Exit Function
    If mRight And mDown Then PressedKey = KeyRightDown: Exit Function
    If mDown And mLeft Then PressedKey = KeyDownLeft: Exit Function
    
    If mUp Then PressedKey = KeyUp: Exit Function
    If mLeft Then PressedKey = KeyLeft: Exit Function
    If mRight Then PressedKey = KeyRight: Exit Function
    If mDown Then PressedKey = KeyDown: Exit Function
    
    
    If (GetAsyncKeystate(vbKey0) And &H1) = &H1 Then
        PressedKey = Key0
        Exit Function
    End If
    If (GetAsyncKeystate(vbKeyNumpad0) And &H1) = &H1 Then
        PressedKey = Key0
        Exit Function
    End If
    If (GetAsyncKeystate(vbKeyInsert) And &H1) = &H1 Then
        PressedKey = Key0
        Exit Function
    End If
    
    If (GetAsyncKeystate(vbKeyReturn) And &H1) = &H1 Then
        PressedKey = KeyEnter
        Exit Function
    End If

End Function
