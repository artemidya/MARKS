Private Sub CommandButton1_Click()
    On Error Resume Next
    Dim file As String
    file = Application.CorelScriptTools.GetFileBox

    
    Dim bytes() As Byte
    Dim fileNumber As Long
    
    fileNumber = FreeFile
    
    Open file For Binary Access Read As fileNumber
        ReDim bytes(LOF(fileNumber) - 1)
        Get fileNumber, , bytes
    Close fileNumber
    
    Dim top As Long, left As Long
    Dim maxCelTop As Long, maxLeft As Long
    Dim residueLeft As Long
    Dim step As Long
    
    step = 16
    maxCelTop = (UBound(bytes) + 1) \ step
    residueLeft = (UBound(bytes) + 1) Mod step
    
    Dim str As String
    
    str = vbTab & "Dim b(" & UBound(bytes) & ") as Byte" & vbCrLf
    For top = 1 To maxCelTop
        str = str & vbTab
        For left = 1 To step - 1
            str = str & "b(" & (top - 1) * step + left - 1 & ")=" & bytes((top - 1) * step + left - 1) & " : "
        Next
        str = str & "b(" & (top - 1) * step + left - 1 & ")=" & bytes((top - 1) * step + left - 1) & vbCrLf
    Next
    
    If residueLeft > 0 Then
        str = str & vbTab
        For left = 1 To residueLeft - 1
            str = str & "b(" & (top - 1) * step + left - 1 & ")=" & bytes((top - 1) * step + left - 1) & " : "
        Next
        str = str & "b(" & (top - 1) * step + left - 1 & ")=" & bytes((top - 1) * step + left - 1)
    End If
    
    TextBox1.text = str
End Sub
