Private Function HoursAndMinets(min As Long) As String
    Dim h As Long
    Dim m As Long
    h = min \ 60
    m = min Mod 60
    If h > 0 Then HoursAndMinets = h & " " & localization.ActiveLng.Item("ActivationForm.hour") & " "
    If m > 0 Then HoursAndMinets = HoursAndMinets & m & " " & localization.ActiveLng.Item("ActivationForm.minutes") & " "
End Function
