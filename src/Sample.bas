Attribute VB_Name = "Sample"
Option Explicit
Public Sub sample1()
    With New GLBChartBuilder
        .HasLine = True
        .HasMarker = False
        .LineWidth = 5
        .LineColor = RGB(50, 150, 50)
        With .CreateFromRange(ActiveSheet.Range("B1:D5000"))
            .Left = 30
            .Top = 25
            .width = 250
            .Height = 250
        End With
    End With
End Sub
Public Sub sample2()
    With New GLBChartBuilder
        .HasLine = False
        .HasMarker = True
        .MarkerSize = 10
        .MarkerColor = RGB(150, 50, 50)
        With .CreateFromRange(ActiveSheet.Range("F1:H5000"))
            .Left = 30
            .Top = 320
            .width = 250
            .Height = 250
        End With
    End With
End Sub
Public Sub sample3()
    With New GLBChartBuilder
        .HasLine = True
        .HasMarker = True
        With .CreateFromRange(ActiveSheet.Range("J1:L5000"))
            .Left = 320
            .Top = 25
            .width = 250
            .Height = 250
        End With
    End With
End Sub
