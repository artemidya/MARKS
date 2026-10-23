Private Function MmPlus() As String
    On Error GoTo e:
    Dim mm As Long
    mm = Int((LastMaxY - Work_Result(CountResults).maximumY) * 25.4)
    If mm > 0 Then MmPlus = " (+" & mm & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
e:
End Function
