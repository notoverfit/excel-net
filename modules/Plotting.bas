Attribute VB_Name = "Plotting"
Public Function InitPlot()
    Dim chartObj As ChartObject
    Dim networks As Worksheet
    
    Set networks = ThisWorkbook.Worksheets("Network")
    
    Set chartObj = networks.ChartObjects.Add(Left:=450, Top:=70, Width:=500, Height:=250)
    With chartObj.Chart
        .ChartType = xlLine
        .SetSourceData networks.Range("G1:H1")
    End With
    
    Set InitPlot = chartObj
End Function

Public Sub UpdatePlot(ByVal plot As ChartObject, ByVal networks As Worksheet, ByVal epoch As Long)
    plot.Chart.SetSourceData networks.Range("G1:H" & epoch + 1)
    DoEvents
End Sub
