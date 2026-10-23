Private Sub Class_Initialize()
    CountProfiles = CountProfiles + 1
    MyProfileNumber = CountProfiles
    'Load
    lFile = VGCore.GMSManager.Projects("PlotCalc").FilePath & "Profile" & MyProfileNumber & ".xml"
    Load
End Sub
