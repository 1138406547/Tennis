VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub RaceOtherDetailCharts()
      Dim MyChart As ChartObject
      On Error Resume Next
      With Sheets("raceOtherDetails")
            .ChartObjects.Delete
            Set MyChart = .ChartObjects.add(left:=500, Top:=40, Width:=310, Height:=300)
            MyChart.Name = "MyChart"
            With MyChart.Chart
                  .ChartType = xlDoughnut   'xl3DColumnStacked
                  .SetSourceData Source:=SetChartSourceData
                  .ChartGroups(1).DoughnutHoleSize = 61.8
                  .FullSeriesCollection(1).ApplyDataLabels
                  .HasTitle = False
                  .HasLegend = False
            End With
            .Activate
      End With
      Set MyChart = Nothing
End Sub


Private Function SetChartSourceData() As Range
      Dim Side%, Way%
      
      With Sheets("raceOtherDetails")
            Side = .Range("L4").Value
            Way = .Range("M4").Value
            If Way = 1 Then
                  If Side = 1 Then
                        Set SetChartSourceData = .Range(.Cells(26, 3), .Cells(29, 3))
                  Else
                        Set SetChartSourceData = .Range(.Cells(26, 4), .Cells(29, 4))
                  End If
            ElseIf Way = 2 Then
                  If Side = 1 Then
                        Set SetChartSourceData = .Range(.Cells(12, 6), .Cells(14, 6))
                  Else
                        Set SetChartSourceData = .Range(.Cells(12, 7), .Cells(14, 7))
                  End If
            Else
                  If Side = 1 Then
                        Set SetChartSourceData = .Range(.Cells(12, 9), .Cells(14, 9))
                  Else
                        Set SetChartSourceData = .Range(.Cells(12, 10), .Cells(14, 10))
                  End If
            End If
      End With
End Function
