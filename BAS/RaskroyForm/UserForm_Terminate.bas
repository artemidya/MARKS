Private Sub UserForm_Terminate()
    'proverit Change
    Break = True
    If UserData.LastProfileSelectionNumber <> Combo1LastSelectListIndex Then
        UserData.LastProfileSelectionNumber = Combo1LastSelectListIndex
        UserData.Save
    End If
    TerminateGDI
    Sleep 200
    BlocksSocket1.Disconnect


    Application.ActiveDocument.Unit = lastUnit
    ActiveDocument.ReferencePoint = lastReferencePoint
End Sub
