VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Sub trainningCharts()
      Dim MyChart As ChartObject
      Dim ChartName$, Xmax%, Ymax%
      On Error Resume Next
      Application.ScreenUpdating = False
      Sheets("trainningCharts").ChartObjects.Delete
      With Sheets("trainningCharts")
            Set MyChart = .ChartObjects.add _
                  (left:=10, _
                  Top:=10, _
                  Width:=800, Height:=400)
            MyChart.Name = "落点与击球点"
            With MyChart.Chart
                  .ChartType = xlXYScatter   'xl3DColumnStacked
                  ChartName = SetChartSourceData
                  .SetSourceData Source:=Sheets("trainningCoordinates").Range(Sheets("trainningCoordinates").Cells(1, 50), Sheets("trainningCoordinates").Cells(10000, 51).End(xlUp))
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
            .Activate
      End With
      Set MyChart = Nothing
      Application.ScreenUpdating = True
End Sub


Private Function SetChartSourceData() As String
      Dim sideChoice$
      Dim dataChoice$
      Dim i%, j%, k%
      Dim rng As Range
      With Sheets("trainningCharts")
            sideChoice = IIf(.Range("R4").Value = 1, "A", "B")
            dataChoice = IIf(.Range("R6").Value = 1, "hit", "drop")
      End With
      
      With Sheets("trainningCoordinates")
            .Columns(50).Clear
            .Columns(51).Clear
            .Cells(1, 50) = 1188
            .Cells(2, 50) = 1188
            .Cells(3, 50) = -1188
            .Cells(4, 50) = -1188
            .Cells(1, 51) = 410
            .Cells(2, 51) = -410
            .Cells(3, 51) = 410
            .Cells(4, 51) = -410
            
            
            If .Cells(1, 1) <> "" Then
                  Dim Side$, hitOrDrop$                           '单元格(1,1)的场地以及模式描述
                  j = InStr(1, .Cells(1, 1), "_")
                  k = InStr(j + 1, .Cells(1, 1), "_")
                  If j = 0 Or k = 0 Then
                        MsgBox "error123"
                        Exit Function
                  End If
                  Side = Mid(.Cells(1, 1).Value, 1, j - 1)                    '显示A还是B的数据
                  hitOrDrop = Mid(.Cells(1, 1).Value, j + 1, k - j - 1)       '球落点还是击球点
                  
                  If (Side = "A方" And sideChoice = "A") Or (Side = "B方" And sideChoice = "B") Then
                        If dataChoice = "hit" Then                            '选择显示击球点
                              If hitOrDrop = "接发球" Then
                                    k = .Cells(10000, 10).End(xlUp).Row
                                    If k >= 3 Then
                                          Set rng = .Range(.Cells(3, 9), .Cells(10000, 10).End(xlUp))
                                          rng.copy
                                          .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                    End If
                              Else
                                    k = .Cells(10000, 2).End(xlUp).Row
                                    If k >= 3 Then
                                          Set rng = .Range(.Cells(3, 1), .Cells(10000, 2).End(xlUp))
                                          rng.copy
                                          .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                    End If
                              End If
                        Else                                                  '选择显示球落点
                              k = .Cells(10000, 4).End(xlUp).Row
                              If k >= 3 Then
                                    Set rng = .Range(.Cells(3, 3), .Cells(10000, 4).End(xlUp))
                                    rng.copy
                                    .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                              End If
                              k = .Cells(10000, 6).End(xlUp).Row
                              If k >= 3 Then
                                    Set rng = .Range(.Cells(3, 5), .Cells(10000, 6).End(xlUp))
                                    rng.copy
                                    .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                              End If
                              k = .Cells(10000, 8).End(xlUp).Row
                              If k >= 3 Then
                                    Set rng = .Range(.Cells(3, 7), .Cells(10000, 8).End(xlUp))
                                    rng.copy
                                    .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                              End If
                        End If
                        SetChartSourceData = Mid(.Cells(1, 1).Value, 1, InStrRev(.Cells(1, 1), "_") - 1)
                  ElseIf Side = "A方" And sideChoice = "B" Then
                        If hitOrDrop = "接发球" Then
                              If .Cells(1, 11).Value <> "" Then
                                    If dataChoice = "hit" Then                      '选择显示击球点
                                          If .Cells(1, 11).Value Like "*接发球*" Then
                                                k = .Cells(10000, 20).End(xlUp).Row
                                                If k >= 3 Then
                                                      Set rng = .Range(.Cells(3, 19), .Cells(10000, 20).End(xlUp))
                                                      rng.copy
                                                      .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                                End If
                                          Else
                                                k = .Cells(10000, 12).End(xlUp).Row
                                                If k >= 3 Then
                                                      Set rng = .Range(.Cells(3, 11), .Cells(10000, 12).End(xlUp))
                                                      rng.copy
                                                      .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                                End If
                                          End If
                                    Else                                            '选择显示球落点
                                          k = .Cells(10000, 14).End(xlUp).Row
                                          If k >= 3 Then
                                                Set rng = .Range(.Cells(3, 13), .Cells(10000, 14).End(xlUp))
                                                rng.copy
                                                .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                          End If
                                          k = .Cells(10000, 16).End(xlUp).Row
                                          If k >= 3 Then
                                                Set rng = .Range(.Cells(3, 15), .Cells(10000, 16).End(xlUp))
                                                rng.copy
                                                .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                          End If
                                          k = .Cells(10000, 18).End(xlUp).Row
                                          If k >= 3 Then
                                                Set rng = .Range(.Cells(3, 17), .Cells(10000, 18).End(xlUp))
                                                rng.copy
                                                .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                          End If
                                    End If
                              End If
                              SetChartSourceData = Mid(.Cells(1, 11).Value, 1, InStrRev(.Cells(1, 11), "_") - 1)
                        Else
                              If .Cells(1, 11).Value <> "" Then
                                    If dataChoice = "hit" Then
                                          If .Cells(1, 9).Value Like "*接发球*" Then
                                                k = .Cells(10000, 18).End(xlUp).Row
                                                If k >= 3 Then
                                                      Set rng = .Range(.Cells(3, 17), .Cells(10000, 18).End(xlUp))
                                                      rng.copy
                                                      .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                                End If
                                          Else
                                                k = .Cells(10000, 10).End(xlUp).Row
                                                If k >= 3 Then
                                                      Set rng = .Range(.Cells(3, 9), .Cells(10000, 10).End(xlUp))
                                                      rng.copy
                                                      .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                                End If
                                          End If
                                    Else
                                          k = .Cells(10000, 14).End(xlUp).Row
                                          If k >= 3 Then
                                                Set rng = .Range(.Cells(3, 13), .Cells(10000, 14).End(xlUp))
                                                rng.copy
                                                .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                          End If
                                          k = .Cells(10000, 16).End(xlUp).Row
                                          If k >= 3 Then
                                                Set rng = .Range(.Cells(3, 15), .Cells(10000, 16).End(xlUp))
                                                rng.copy
                                                .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                          End If
                                          k = .Cells(10000, 12).End(xlUp).Row
                                          If k >= 3 Then
                                                Set rng = .Range(.Cells(3, 11), .Cells(10000, 12).End(xlUp))
                                                rng.copy
                                                .Cells(10000, 50).End(xlUp).Offset(1).PasteSpecial
                                          End If
                                    End If
                              End If
                              SetChartSourceData = Mid(.Cells(1, 9).Value, 1, InStrRev(.Cells(1, 9), "_") - 1)
                        End If
                  Else
                        Exit Function
                  End If
            End If
            Application.CutCopyMode = False
      End With
      Application.Goto Sheets("trainningCoordinates").Range("a1")
End Function




