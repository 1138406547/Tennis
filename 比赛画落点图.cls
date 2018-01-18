VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub raceCharts()
      Dim MyChart As ChartObject
      Dim ChartName$, Xmax%, Ymax%
      On Error Resume Next
      Sheets("raceCharts").ChartObjects.Delete
      With Sheets("raceCharts")
            Set MyChart = .ChartObjects.add _
                  (left:=10, _
                  Top:=10, _
                  Width:=800, Height:=400)
            MyChart.Name = "发球落点与击球点统计"
            With MyChart.Chart
                  .ChartType = xlXYScatter   'xl3DColumnStacked
                  ChartName = SetChartSourceData
                  .SetSourceData Source:=Sheets("raceCoordinates").Range(Sheets("raceCoordinates").Cells(1, 24), Sheets("raceCoordinates").Cells(10000, 25).End(xlUp))
                  .HasTitle = True
                  .HasLegend = False
                  Xmax = IIf(Sheets("raceCharts").Range("R24") = "", 1280, Sheets("raceCharts").Range("R24"))
                  Ymax = IIf(Sheets("raceCharts").Range("R25") = "", 410, Sheets("raceCharts").Range("R25"))
                  .Axes(xlCategory).MinimumScale = -Xmax
                  .Axes(xlCategory).MaximumScale = Xmax
                  .Axes(xlCategory).MajorUnit = Xmax / 4
                  .Axes(xlCategory).ReversePlotOrder = True
                  .Axes(xlValue).MaximumScale = Ymax
                  .Axes(xlValue).MinimumScale = -Ymax
                  .Axes(xlValue).MajorUnit = Ymax / 4
                  .Axes(xlValue).ReversePlotOrder = True
                  With .ChartTitle
                        .Text = ChartName
                        .Font.Name = "宋体"
                        .Font.Size = 12
                  End With
            End With
      End With
      Sheets("raceCharts").Select
      Set MyChart = Nothing
End Sub


Private Function SetChartSourceData() As String
      Dim Side%, hitOrServe%
      Dim isACE As Boolean, isFirstServe As Boolean, isSecondServe As Boolean
      Dim rng As Range, tmp As Range
      Dim i%
      With Sheets("raceCharts")
            Side = .Range("R4").Value
            hitOrServe = .Range("R6").Value
            isACE = .Range("R11").Value
            isFirstServe = .Range("R12").Value
            isSecondServe = .Range("R13").Value
      End With
      
      With Sheets("raceCoordinates")
            .Columns("x:y").Clear
            .Cells(1, 24) = 1188
            .Cells(2, 24) = 1188
            .Cells(3, 24) = -1188
            .Cells(4, 24) = -1188
            .Cells(1, 25) = 410
            .Cells(2, 25) = -410
            .Cells(3, 25) = 410
            .Cells(4, 25) = -410
            If hitOrServe = 1 Then                    '击球点
                  If Side = 1 Then                    'A方
                        If .Cells(2, 17) <> "" Then
                              Set rng = .Range(.Cells(2, 17), .Cells(10000, 18).End(xlUp))
                              rng.copy
                              i = .Cells(10000, 24).End(xlUp).Row
                              .Cells(i + 1, 24).PasteSpecial
                        End If
                        SetChartSourceData = "A方击球点统计"
                  Else                                'B方
                        If .Cells(2, 19) <> "" Then
                              Set rng = .Range(.Cells(2, 19), .Cells(10000, 20).End(xlUp))
                              rng.copy
                              i = .Cells(10000, 24).End(xlUp).Row
                              .Cells(i + 1, 24).PasteSpecial
                        End If
                        SetChartSourceData = "B方击球点统计"
                  End If
            Else                                                  '发球落点
                  If Side = 1 Then                                'A方
                        If isACE Then                             'ACE被选择
                              If .Cells(2, 1) <> "" Then          'A一发ACE
                                    Set rng = .Range(.Cells(2, 1), .Cells(10000, 2).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              If .Cells(2, 5) <> "" Then          'A二发ACE
                                    Set rng = .Range(.Cells(2, 5), .Cells(10000, 6).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              SetChartSourceData = "A方ACE"
                        End If
                        If isFirstServe Then
                              If .Cells(2, 3) <> "" Then          'A一发非ACE
                                    Set rng = .Range(.Cells(2, 3), .Cells(10000, 4).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              If .Cells(2, 1) <> "" Then          'A一发ACE
                                    Set rng = .Range(.Cells(2, 1), .Cells(10000, 2).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              SetChartSourceData = IIf(SetChartSourceData = "", "A方一发落点", SetChartSourceData & " + A方一发落点")
                        End If
                        If isSecondServe Then                     '二发被选择
                              If .Cells(2, 5) <> "" Then          'A二发ACE
                                    Set rng = .Range(.Cells(2, 5), .Cells(10000, 6).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              If .Cells(2, 7) <> "" Then          'A二发非ACE
                                    Set rng = .Range(.Cells(2, 7), .Cells(10000, 8).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              SetChartSourceData = IIf(SetChartSourceData = "", "A方二发落点", SetChartSourceData & " + A方二发落点")
                        End If
                  ElseIf Side = 2 Then                            'B方
                        If isACE Then                             'ACE被选择
                              If .Cells(2, 9) <> "" Then          'B一发ACE
                                    Set rng = .Range(.Cells(2, 9), .Cells(10000, 10).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              If .Cells(2, 13) <> "" Then          'B二发ACE
                                    Set rng = .Range(.Cells(2, 13), .Cells(10000, 14).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              SetChartSourceData = "B方ACE"
                        End If
                        If isFirstServe Then
                              If .Cells(2, 11) <> "" Then          'B一发非ACE
                                    Set rng = .Range(.Cells(2, 11), .Cells(10000, 12).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              If .Cells(2, 9) <> "" Then          'B一发ACE
                                    Set rng = .Range(.Cells(2, 9), .Cells(10000, 10).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              SetChartSourceData = IIf(SetChartSourceData = "", "B方一发落点", SetChartSourceData & " + B方一发落点")
                        End If
                        If isSecondServe Then          '二发被选择
                              If .Cells(2, 13) <> "" Then          'B二发ACE
                                    Set rng = .Range(.Cells(2, 13), .Cells(10000, 14).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              If .Cells(2, 15) <> "" Then          'B二发非ACE
                                    Set rng = .Range(.Cells(2, 15), .Cells(10000, 16).End(xlUp))
                                    rng.copy
                                    i = .Cells(10000, 24).End(xlUp).Row
                                    .Cells(i + 1, 24).PasteSpecial
                              End If
                              SetChartSourceData = IIf(SetChartSourceData = "", "B方二发落点", SetChartSourceData & " + B方二发落点")
                        End If
                  End If
            End If
            Set rng = .Range(.Cells(1, 24), .Cells(10000, 24).End(xlUp))
            For Each tmp In rng
                  If SetChartSourceData Like "A*" Then
                        If SetChartSourceData Like "*击球点*" Then
                              If tmp.Value < 0 Then
                                    tmp.Value = tmp.Value * (-1)
                                    tmp.Offset(, 1).Value = tmp.Offset(, 1).Value * (-1)
                              End If
                        Else
                              If tmp.Value > 0 Then
                                    tmp.Value = tmp.Value * (-1)
                                    tmp.Offset(, 1).Value = tmp.Offset(, 1).Value * (-1)
                              End If
                        End If
                  Else
                        If SetChartSourceData Like "*击球点*" Then
                              If tmp.Value > 0 Then
                                    tmp.Value = tmp.Value * (-1)
                                    tmp.Offset(, 1).Value = tmp.Offset(, 1).Value * (-1)
                              End If
                        Else
                              If tmp.Value < 0 Then
                                    tmp.Value = tmp.Value * (-1)
                                    tmp.Offset(, 1).Value = tmp.Offset(, 1).Value * (-1)
                              End If
                        End If
                  End If
            Next
      End With
End Function


