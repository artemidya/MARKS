Private Sub UserForm_Initialize()
    UserData.Load
    ListBox1.ColumnCount = 2
   
    Dim q As Long
    For q = 0 To ActiveLng.count - 1
        With ListBox1
            .AddItem
            .List(q, 0) = ActiveLng.Key(q)
            .List(q, 1) = ActiveLng.value(q)
        End With
    Next
    ListBox1.ListIndex = 0
    Label7.Caption = "ISO639=" & LCase(localization.CodeISO639)
End Sub
