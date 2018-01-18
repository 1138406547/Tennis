VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const WAITING As Integer = 19

Sub judge()

      On Error GoTo errro
      
      Dim i%, ii%, j%, k%
      Dim xPos%, yPos%, zPos%, statPos%, speedPos%, seqPos%
      Dim filenumber%
      Dim fileDestination$
      fileDestination = Sheets("main").Range("M18").Value & "\" & Sheets("main").Range("M19").Value & ".txt"
      If Dir(fileDestination, vbNormal) = "" Then
            fileDestination = ThisWorkbook.Path & "\points.txt"
      End If
      
      Dim tempLine$, exchangeStr$
      exchangeStr = Sheets("main").Range("Q17").Value
      
      Dim code%, serveFlag%, exchangeFlag%, latestHitStat%
      exchangeFlag = 1
      
      Dim mode%, ASideTrainningMode%, BSideTrainningMode%
      'mode: 比赛(1)；A方单人训练(2);B方单人训练(3);双方同时训练(4)
      'ASideTrainningMode | BSideTrainningMode: A或者B方训练模式：发球(1);接发球(2);击球(3)
      mode = Sheets("main").Range("L3").Value
      ASideTrainningMode = Sheets("main").Range("M6").Value
      BSideTrainningMode = Sheets("main").Range("N6").Value
      
      Dim trainningSide%, trainningLevel%
      'trainningSide: 训练者所在场地方位：A(1);B(-1);A and B(0)
      'trainningLevel:训练难度：易(1);中(2);难(3)
      If mode = 2 Then
            trainningSide = 1
      ElseIf mode = 3 Then
            trainningSide = -1
      ElseIf mode = 4 Then
            trainningSide = 0
      End If
      trainningLevel = Sheets("main").Range("P3").Value
            
      
      Dim last As New ball
      Dim current As New ball
      Dim useful As New ball
      
      Dim currentClone As New ball
      Dim lastClone As New ball
      Dim usefulClone As New ball
      
      Dim result As New score
      
      Dim last_hit%, last_successful_hit%
      
      Dim file, line
      
      i = 1
      ii = 1
      last.stat = WAITING
      lastClone.stat = WAITING
      
      
      Sheets("main").Activate
      ActiveSheet.Range(Columns(1), Columns(11)).Clear
      
     
      filenumber = FreeFile
      Open fileDestination For Input As #filenumber
            file = Split(StrConv(InputB(LOF(filenumber), filenumber), vbUnicode), Chr(10))
      Close #filenumber
      
      For Each line In file
            
            tempLine = Replace(line, Chr(9), Chr(32))
            tempLine = Replace(tempLine, ",", Chr(32))
            tempLine = Application.WorksheetFunction.Trim(tempLine)
            If tempLine = "" Then
            Else
                  xPos = InStr(1, tempLine, " ")
                  yPos = InStr(xPos + 1, tempLine, " ")
                  zPos = InStr(yPos + 1, tempLine, " ")
                  statPos = InStr(zPos + 1, tempLine, " ")
                  speedPos = InStr(statPos + 1, tempLine, " ")
                  seqPos = InStr(speedPos + 1, tempLine, " ")
                  With ActiveSheet
                        If statPos = 0 Then
                              current.stat = --Mid(tempLine, zPos + 1, Len(tempLine) - zPos)
                        Else
                              current.stat = --Mid(tempLine, zPos + 1, statPos - zPos - 1)
                        End If
                        If current.stat <> 0 Then
                              current.x = --Mid(tempLine, 1, xPos - 1)
                              current.y = --Mid(tempLine, xPos + 1, yPos - xPos - 1)
                              current.z = --Mid(tempLine, yPos + 1, zPos - yPos - 1)
                              If statPos <> 0 Then
                                    If speedPos = 0 Then
                                          current.speed = --Mid(tempLine, statPos + 1, Len(tempLine) - statPos)
                                    Else
                                          current.speed = --Mid(tempLine, statPos + 1, speedPos - statPos - 1)
                                          If seqPos <> 0 Then
                                                current.timeStamp = Mid(tempLine, seqPos + 1, Len(tempLine) - seqPos)
                                          End If
                                    End If
                              End If
                              .Cells(i, 1) = current.x
                              .Cells(i, 2) = current.y
                              .Cells(i, 3) = current.z
                              .Cells(i, 4) = current.stat
                              .Cells(i, 5) = current.speed
                              If mode = 1 Then
                                    Call match(current, last, useful, code, i, serveFlag, exchangeFlag, exchangeStr, latestHitStat, result, last_hit, last_successful_hit)
                                    i = i + 1
                              Else
                                    If trainningSide <> 0 Then
                                          Call JudgeTrainning(current, last, useful, i, IIf(trainningSide = 1, ASideTrainningMode, BSideTrainningMode), trainningSide)
                                    Else
                                          Call currentClone.clone(current)
                                          Call JudgeTrainning(current, last, useful, i, ASideTrainningMode, 1)
                                          Call JudgeTrainning(currentClone, lastClone, usefulClone, ii, BSideTrainningMode, -1)
                                    End If
                                    i = i + 1
                                    ii = ii + 1
                              End If
                        End If
                  End With
            End If
      Next
      
      Sheets("trainningOtherDetails").Range("E7:F14").Clear
      
      If mode = 1 Then
            Call result.showBout(Sheets("main").Cells(i + 2, 2))
            Call result.showGame(Sheets("main").Cells(i + 2, 4))
            Call result.showSet(Sheets("main").Cells(i + 2, 6))
            Call raceAnalysis
      ElseIf mode = 2 Then
            Sheets("trainningCoordinates").Cells.Clear
            Call TrainningAnalysis(6, ASideTrainningMode, trainningSide, trainningLevel)
      ElseIf mode = 3 Then
            Sheets("trainningCoordinates").Cells.Clear
            Call TrainningAnalysis(9, BSideTrainningMode, trainningSide, trainningLevel)
      Else
            Sheets("trainningCoordinates").Cells.Clear
            Call TrainningAnalysis(6, ASideTrainningMode, 1, trainningLevel)
            Call TrainningAnalysis(9, BSideTrainningMode, -1, trainningLevel)
      End If
      
errro:
      Close #filenumber
      
End Sub


