Property Get CropBMPoffset() As Long
    On Error GoTo 10
    CropBMPoffset = CLng(GetSetting(AppName, "CropBMP", "CropBMPoffset"))
    If CropBMPoffset < 0 Then CropBMPoffset = 0
10:
End Property
