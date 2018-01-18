Attribute VB_Name = "模块3"
Option Explicit

Const CMSINGLE As Integer = 411
Const CMDOUBLE As Integer = 548
Const CMSERVE As Integer = 640
Const CMBASE As Integer = 1188
Const CMINOROUT As Integer = 205
Const CMHARD As Integer = 914
Const CMOUT As Integer = 2000

Const MMSINGLE As Integer = 4115
Const MMDOUBLE As Integer = 5487
Const MMSERVE As Integer = 6401
Const MMBASE As Integer = 11887
Const MMINOROUT As Integer = 2057
Const MMHARD As Integer = 9144
Const MMOUT As Integer = 20000

Const WAITING As Integer = 19
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''比赛逻辑部分'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub match(current As ball, last As ball, useful As ball, code%, i%, serveFlag%, exchangeFlag%, exchangeStr$, latestHitStat%, result As score, last_hit%, last_successful_hit%)
      Dim edge%, baseLine%, serveLine%, outOfSight%, serve_edge%
      Dim commit$
      Dim raceMode%
      
      raceMode = IIf(Sheets("main").Range("Q15").Value = "", 6, Sheets("main").Range("Q15").Value)
      
      commit = "waitingForNextServe"
      If Sheets("main").Range("N3").Value = 1 Then
            If ActiveSheet.Range("M3").Value = 1 Then
                  edge = CMSINGLE
            Else
            
                  edge = CMDOUBLE
            End If
            baseLine = CMBASE
            serveLine = CMSERVE
            outOfSight = CMOUT
            serve_edge = CMSINGLE
      Else
            If ActiveSheet.Range("M3").Value = 1 Then
                  edge = MMSINGLE
            Else
                  edge = MMDOUBLE
            End If
            baseLine = MMBASE
            serveLine = MMSERVE
            outOfSight = MMOUT
            serve_edge = MMSINGLE
      End If

      With Sheets("main")
            
            
            Select Case current.stat
            
            Case 1
                  Select Case last.stat
                  Case WAITING
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                  Case 1, 10
                        If Sgn(last.x) <> Sgn(current.x) And Sgn(last.x) <> 0 Then
                              commit = "hitBack"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last_successful_hit = last_hit
                              Call last.clone(current)
                              last_hit = Sgn(current.x)
                        Else
                              commit = "hitBackTwice,boutEnd"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              last_hit = Sgn(current.x)
                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                              Call initialize(last_successful_hit, last_hit, useful)
                        End If
                  Case 2
                        commit = "Error Code 21"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        code = 21
                        last.stat = WAITING
                        serveFlag = 0
                        Call initialize(last_successful_hit, last_hit, useful)
                  Case 3
                        If Sgn(last.x) = Sgn(current.x) Then
                              If Abs(latestHitStat) = 2 Then
                                    commit = "return"
                              Else
                                    commit = "hitBack"
                              End If
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              Call last.clone(current)
                              last_hit = Sgn(current.x)
                        Else
                              commit = "Error! Code 31"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              code = 31
                              last.stat = WAITING
                              Call initialize(last_successful_hit, last_hit, useful)
                        End If
                  Case 4, 9
                        Select Case useful.stat
                        Case 1, 10
                              If Sgn(useful.x) <> Sgn(current.x) And Sgn(useful.x) <> 0 Then
                                    commit = "hitBack"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last_successful_hit = last_hit
                                    Call last.clone(current)
                                    last_hit = Sgn(current.x)
                              Else
                                    commit = "hitBackTwice,boutEnd"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    last_hit = Sgn(current.x)
                                    Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                    Call initialize(last_successful_hit, last_hit, useful)
                              End If
                        Case 2
                              commit = "Error Code 21"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              code = 21
                              last.stat = WAITING
                              serveFlag = 0
                              Call initialize(last_successful_hit, last_hit, useful)
                        Case Else
                              commit = "LogicError"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call initialize(last_successful_hit, last_hit, useful)
                        End Select
                  Case 6
                        commit = "Error! Code 61"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        code = 61
                        last.stat = WAITING
                        Call initialize(last_successful_hit, last_hit, useful)
                  End Select
                  latestHitStat = 1
            
            Case 2
                  If last.stat = WAITING Then
                        If serveFlag = 0 Then
                              commit = "firstServe"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last_hit = Sgn(current.x)
                              Call last.clone(current)
                              serveFlag = serveFlag + 1
                        ElseIf serveFlag = 1 Then
                              commit = "secondServe"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last_hit = Sgn(current.x)
                              Call last.clone(current)
                              serveFlag = serveFlag + 1
                        Else
                              commit = "serveFlagErrorAtWaitingFor2"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              Exit Sub
                        End If
                  Else
                        commit = "Error! Code " & last.stat & "2"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        code = last.stat * 10 + 2
                        last.stat = WAITING
                        Call initialize(last_successful_hit, last_hit, useful)
                  End If
                  latestHitStat = 2
            
            Case 3
                  Select Case last.stat
                  Case WAITING
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                  Case 1, 10
                        If Abs(current.x) > baseLine Or Abs(current.y) > edge Then
                              commit = "outOfCourt,boutEnd"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                              Call initialize(last_successful_hit, last_hit, useful)
                        Else
                              If Sgn(last.x) = Sgn(current.x) Then
                                    commit = "hitBackFault,boutEnd"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                    Call initialize(last_successful_hit, last_hit, useful)
                              Else
                                    commit = "in"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last_successful_hit = last_hit
                                    Call last.clone(current)
                              End If
                        End If
                  Case 2
                        If Sgn(current.x) <> Sgn(last.x) And Sgn(current.y) * last.y <= 0 And Abs(current.x) <= serveLine And Abs(current.y) <= serve_edge Then
                              If serveFlag = 2 Then
                                    commit = "secondServeIn"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    serveFlag = 0
                                    last_successful_hit = last_hit
                                    Call last.clone(current)
                              Else
                                    commit = "firstServeIn"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    serveFlag = 0
                                    last_successful_hit = last_hit
                                    Call last.clone(current)
                              End If
                        Else
                              If serveFlag = 1 Then
                                    commit = "fault,waitingForSecondServe"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    Call initialize(last_successful_hit, last_hit, useful)
                              Else
                                    commit = "doubleFault,boutEnd"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    serveFlag = 0
                                    Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                    Call initialize(last_successful_hit, last_hit, useful)
                              End If
                        End If
                  Case 3
                        If Sgn(last.x) = Sgn(current.x) Then
                              commit = "touchDownTwice,boutEnd"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                              Call initialize(last_successful_hit, last_hit, useful)
                        Else
                              commit = "Error! Code 33"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              code = 33
                              last.stat = WAITING
                              Call initialize(last_successful_hit, last_hit, useful)
                        End If
                  Case 4
                        If useful.stat = 2 Then
                              If Sgn(current.x) * useful.x <= 0 And Sgn(current.y) * useful.y <= 0 And Abs(current.x) <= serveLine And Abs(current.y) <= serve_edge Then
                                    If serveFlag = 2 Then
                                          commit = "let,secondServe"
                                          Call description(.Cells(i, 6), commit, exchangeFlag)
                                          last.stat = WAITING
                                          serveFlag = serveFlag - 1
                                          Call initialize(last_successful_hit, last_hit, useful)
                                    Else
                                          commit = "let,firstServe"
                                          Call description(.Cells(i, 6), commit, exchangeFlag)
                                          last.stat = WAITING
                                          serveFlag = serveFlag - 1
                                          Call initialize(last_successful_hit, last_hit, useful)
                                    End If
                              Else
                                    If serveFlag = 1 Then
                                          commit = "fault,waitingForSecondServe"
                                          Call description(.Cells(i, 6), commit, exchangeFlag)
                                          last.stat = WAITING
                                          Call initialize(last_successful_hit, last_hit, useful)
                                    Else
                                          commit = "doubleFault,boutEnd"
                                          Call description(.Cells(i, 6), commit, exchangeFlag)
                                          last.stat = WAITING
                                          serveFlag = 0
                                          Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                          Call initialize(last_successful_hit, last_hit, useful)
                                    End If
                              End If
                        ElseIf useful.stat = 1 Then
                              If Sgn(useful.x) <> Sgn(current.x) And Abs(current.x) <= baseLine And Abs(current.y) <= edge Then
                                    commit = "in"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last_successful_hit = last_hit
                                    Call last.clone(current)
                              Else
                                    commit = "hitBackNetDown,boutEnd"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                    Call initialize(last_successful_hit, last_hit, useful)
                              End If
                        Else
                              commit = "LogicError"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call initialize(last_successful_hit, last_hit, useful)
                        End If
                  Case 6
                        commit = "Error! Code 63"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        code = 63
                        last.stat = WAITING
                        Call initialize(last_successful_hit, last_hit, useful)
                  Case 9
                        If useful.stat = 2 Then
                              If Sgn(current.x) * useful.x <= 0 And Sgn(current.y) * useful.y <= 0 And Abs(current.x) <= serveLine And Abs(current.y) <= serve_edge Then
                                    If serveFlag = 2 Then
                                          commit = "secondServeIn"
                                          Call description(.Cells(i, 6), commit, exchangeFlag)
                                          last_successful_hit = last_hit
                                          Call last.clone(current)
                                          serveFlag = 0
                                    Else
                                          commit = "firstServeIn"
                                          Call description(.Cells(i, 6), commit, exchangeFlag)
                                          last_successful_hit = last_hit
                                          Call last.clone(current)
                                          serveFlag = 0
                                    End If
                              Else
                                    If serveFlag = 1 Then
                                          commit = "fault,waitingForSecondServe"
                                          Call description(.Cells(i, 6), commit, exchangeFlag)
                                          last.stat = WAITING
                                          Call initialize(last_successful_hit, last_hit, useful)
                                    Else
                                          commit = "doubleFault,boutEnd"
                                          Call description(.Cells(i, 6), commit, exchangeFlag)
                                          last.stat = WAITING
                                          Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                          Call initialize(last_successful_hit, last_hit, useful)
                                          serveFlag = 0
                                    End If
                              End If
                        ElseIf useful.stat = 1 Then
                              If Sgn(useful.x) <> Sgn(current.x) And Abs(current.x) <= baseLine And Abs(current.y) <= edge Then
                                    commit = "in"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last_successful_hit = last_hit
                                    Call last.clone(current)
                              Else
                                    commit = "hitBackFault,boutEnd"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                    Call initialize(last_successful_hit, last_hit, useful)
                              End If
                        Else
                              commit = "LogicError"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call initialize(last_successful_hit, last_hit, useful)
                        End If
                  End Select
            
            
            Case 4
                  Select Case last.stat
                  Case WAITING
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                  Case 1, 2, 10
                        commit = "touchNet"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        Call useful.clone(last)
                        If useful.stat = 10 Then useful.stat = 1
                        Call last.clone(current)
                  Case 3, 4, 6
                        commit = "Error! Code " & last.stat & "4"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        code = last.stat * 10 + 4
                        last.stat = WAITING
                        Call initialize(last_successful_hit, last_hit, useful)
                  Case 9
                        commit = "touchNet"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        Call last.clone(current)
                  End Select
            
            
            Case 6
                  Select Case last.stat
                  Case WAITING
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                  Case 1, 10
                        If Abs(current.x) > baseLine Or Abs(current.y) > edge Then
                              commit = "outGuess,boutEnd"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                              Call initialize(last_successful_hit, last_hit, useful)
                        Else
                              commit = "outOfSight"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              Call useful.clone(last)
                              If useful.stat = 10 Then useful.stat = 1
                              Call last.clone(current)
                        End If
                  Case 2
                        If Abs(current.x) > baseLine Or Abs(current.y) > edge Then
                              If serveFlag = 2 Then
                                    commit = "doubleFault,boutEnd"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    serveFlag = 0
                                    Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                    Call initialize(last_successful_hit, last_hit, useful)
                              Else
                                    commit = "fault,waitingForSecondServe"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    Call initialize(last_successful_hit, last_hit, useful)
                              End If
                        Else
                              commit = "outOfSight"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              Call useful.clone(last)
                              Call last.clone(current)
                        End If
                  Case 3
                        If Abs(current.x) = outOfSight Then
                              commit = "outOfSIght,boutEnd"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                              Call initialize(last_successful_hit, last_hit, useful)
                        Else
                              commit = "outOfSight"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              Call useful.clone(last)
                              Call last.clone(current)
                        End If
                  Case 4, 9
'                        If Abs(current.x) = outofsight Then
'                              call description(.cells(i,6),"outOfSIght,boutEnd"
'
'                              last.stat = WAITING
'                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr,racemode)
'                              Call initialize(last_successful_hit, last_hit,serveFlag, useful)
'                        Else
'                              call description(.cells(i,6),"outOfSight"
'                              last.stat = 6
'                        End If
                        If useful.stat = 2 Then
                              If serveFlag = 2 Then
                                    commit = "outGuess,boutEnd"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    serveFlag = 0
                                    last.stat = WAITING
                                    Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                    Call initialize(last_successful_hit, last_hit, useful)
                              Else
                                    commit = "faultGuess,waitingForSecondServe"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    Call initialize(last_successful_hit, last_hit, useful)
                              End If
                        Else
                              commit = "outGuess,boutEnd"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                              Call initialize(last_successful_hit, last_hit, useful)
                        End If
                  Case 6
                        If Abs(current.x) = outOfSight Then
                              commit = "outOfSIght,boutEnd"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                              Call initialize(last_successful_hit, last_hit, useful)
                        Else
                              commit = "Error! Code 66"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              code = 66
                              last.stat = WAITING
                              Call initialize(last_successful_hit, last_hit, useful)
                        End If
                  Case 10
                        If Abs(current.x) = outOfSight Then
                              commit = "outOfSIght,boutEnd"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              last.stat = WAITING
                              Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                              Call initialize(last_successful_hit, last_hit, useful)
                        Else
                              commit = "outOfSight"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              Call useful.clone(last)
                              Call last.clone(current)
                              If useful.stat = 10 Then useful.stat = 1
                        End If
                  End Select
            
            
            Case 9
                  If last.stat = WAITING Then
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                  ElseIf last.stat = 6 Then
                        commit = "backInSight"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        Call last.clone(current)
                  Else
                        commit = "Error! Code " & last.stat & "9"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        code = last.stat * 10 + 9
                        last.stat = WAITING
                        Call initialize(last_successful_hit, last_hit, useful)
                  End If
            
            
            Case 10
                  If last.stat = WAITING Then
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                  ElseIf last.stat = 6 Then
                        Select Case useful.stat
                        Case 1
                              If Sgn(useful.x) <> Sgn(current.x) Then
                                    commit = "hitBackGuess"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last_successful_hit = Sgn(useful.x)
                                    last_hit = Sgn(current.x)
                                    Call last.clone(current)
                              Else
                                    commit = "hitBackTwice,boutEnd"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last.stat = WAITING
                                    Call result.score(last_successful_hit, last_hit, exchangeFlag, exchangeStr, raceMode)
                                    Call initialize(last_successful_hit, last_hit, useful)
                              End If
                        Case 2
                              commit = "Error Code 21"
                              Call description(.Cells(i, 6), commit, exchangeFlag)
                              code = 21
                              last.stat = WAITING
                              Call initialize(last_successful_hit, last_hit, useful)
                        Case 3
                              If Sgn(current.x) = Sgn(useful.x) Then
                                    If Abs(latestHitStat) = 2 Then
                                          commit = "return"
                                    Else
                                          commit = "hitBackGuess"
                                    End If
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    last_hit = Sgn(current.x)
                                    Call last.clone(current)
                              Else
                                    commit = "Error! Code 31"
                                    Call description(.Cells(i, 6), commit, exchangeFlag)
                                    code = 31
                                    last.stat = WAITING
                                    Call initialize(last_successful_hit, last_hit, useful)
                              End If
                        End Select
                  Else
                        commit = "Error! Code " & last.stat & "10"
                        Call description(.Cells(i, 6), commit, exchangeFlag)
                        last.stat = WAITING
                        code = last.stat * 100 + 10
                        Call initialize(last_successful_hit, last_hit, useful)
                  End If
                  latestHitStat = 1
            End Select
            
            
      End With


End Sub

Private Sub initialize(last_successful_hit%, last_hit%, useful As ball)
        last_successful_hit = 0
        last_hit = 0
        useful.stat = WAITING
        useful.x = 0
        useful.y = 0
End Sub


Private Sub description(rng As Range, words As String, exchangeFlag%)
      rng.Value = words
      If words Like "*boutEnd" Then
            rng.Interior.ColorIndex = 40
      End If
      If Not words Like "waitingForNextServe" Then
            rng.Offset(, 1).Value = exchangeFlag
      End If

End Sub


