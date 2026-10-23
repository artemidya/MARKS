Private Sub SetIcons()
    'SetIcons
    Dim q As Long
    On Error GoTo 10
    CreateFile_14_bmp
    CreateFile_0_bmp
    CreateFile_0_5_bmp
    CreateFile_1_bmp
    CreateFile_2_bmp
    CreateFile_3_bmp
    CreateFile_4_bmp
    CreateFile_5_bmp
    CreateFile_6_bmp
    CreateFile_7_bmp
    CreateFile_9_bmp
    CreateFile_10_bmp
    CreateFile_12_bmp
    CreateFile_11_bmp
    CreateFile_8_bmp
    CreateFile_13_bmp
    With Application.CommandBars("PlotCalc")
        .Controls(1).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "14.bmp"
        .Controls(2).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "0.bmp"
        .Controls(3).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "0_5.bmp"
        .Controls(4).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "1.bmp"
        .Controls(5).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "2.bmp"
        .Controls(6).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "3.bmp"
        .Controls(7).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "4.bmp"
        .Controls(8).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "5.bmp"
        .Controls(9).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "6.bmp"
        .Controls(10).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "7.bmp"
        .Controls(11).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "9.bmp"
        .Controls(12).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "10.bmp"
        .Controls(13).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "13.bmp"
        .Controls(14).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "12.bmp"
        .Controls(15).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "11.bmp"
        .Controls(16).SetCustomIcon VGCore.GMSManager.Projects("PlotCalc").FilePath & "8.bmp"
    End With
    'DeletFiles
    
    With VGCore.GMSManager.Projects("PlotCalc")
         KillFile .FilePath & "14.bmp"
         KillFile .FilePath & "0.bmp"
         KillFile .FilePath & "0_5.bmp"
         KillFile .FilePath & "1.bmp"
         KillFile .FilePath & "2.bmp"
         KillFile .FilePath & "3.bmp"
         KillFile .FilePath & "4.bmp"
         KillFile .FilePath & "5.bmp"
         KillFile .FilePath & "6.bmp"
         KillFile .FilePath & "7.bmp"
         KillFile .FilePath & "9.bmp"
         KillFile .FilePath & "10.bmp"
         KillFile .FilePath & "11.bmp"
         KillFile .FilePath & "8.bmp"
         KillFile .FilePath & "12.bmp"
         KillFile .FilePath & "13.bmp"
    End With
10:
End Sub
