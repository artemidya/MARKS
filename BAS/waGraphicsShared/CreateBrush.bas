Public Function CreateBrush(Optional ColorRGB As Long = 6579300, Optional Opacity As Long = 100) As waBrush
    Set CreateBrush = New waBrush
    CreateBrush.SetColor ColorRGB, Opacity
End Function
