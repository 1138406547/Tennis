VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "score"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private A_bout() As Integer
Private A_game() As Integer
Private A_set As Integer

Private B_bout() As Integer
Private B_game() As Integer
Private B_set As Integer

Private boutCount As Integer        '局序号
Private gameCount As Integer        '盘序号

Private Sub Class_Initialize()
      ReDim A_bout(50)
      ReDim B_bout(50)
      ReDim A_game(10)
      ReDim B_game(10)
      boutCount = 0
      gameCount = 0
End Sub


Public Sub score(last_successful_hit%, last_hit%, exchangeFlag%, exchangeStr$, Optional gameMode% = 6)
      If last_successful_hit <> 0 Then
            If last_successful_hit * exchangeFlag = 1 Then
                  A_bout(boutCount) = A_bout(boutCount) + 1
            Else
                  B_bout(boutCount) = B_bout(boutCount) + 1
            End If
      ElseIf last_hit * exchangeFlag = 1 Then
            B_bout(boutCount) = B_bout(boutCount) + 1
      Else
            A_bout(boutCount) = A_bout(boutCount) + 1
      End If
      If gameMode = 7 Then
单抢七局:
            If A_bout(boutCount) >= 7 And A_bout(boutCount) >= B_bout(boutCount) + 2 Then
                  A_set = A_set + 1
                  A_game(gameCount) = A_game(gameCount) + 1
                  gameCount = gameCount + 1
                  If gameCount > UBound(A_game) Then
                        ReDim Preserve A_game(gameCount + 10)
                        ReDim Preserve B_game(gameCount + 10)
                  End If
                  boutCount = boutCount + 1
                  If boutCount > UBound(A_bout) Then
                        ReDim Preserve A_bout(boutCount + 50)
                        ReDim Preserve B_bout(boutCount + 50)
                  End If
                  If strExam(exchangeStr) = "1" Then
                        exchangeFlag = exchangeFlag * (-1)
                  End If
            ElseIf B_bout(boutCount) >= 7 And B_bout(boutCount) >= A_bout(boutCount) + 2 Then
                  B_set = B_set + 1
                  B_game(gameCount) = B_game(gameCount) + 1
                  gameCount = gameCount + 1
                  If gameCount > UBound(A_game) Then
                        ReDim Preserve A_game(gameCount + 5)
                        ReDim Preserve B_game(gameCount + 5)
                  End If
                  boutCount = boutCount + 1
                  If boutCount > UBound(A_bout) Then
                        ReDim Preserve A_bout(boutCount + 50)
                        ReDim Preserve B_bout(boutCount + 50)
                  End If
                  If strExam(exchangeStr) = "1" Then
                        exchangeFlag = exchangeFlag * (-1)
                  End If
            End If
      ElseIf gameMode = 4 Or gameMode = 6 Then
四局or六局抢七:
            If A_game(gameCount) = gameMode And B_game(gameCount) = A_game(gameCount) Then
抢七局:
                  If A_bout(boutCount) >= 7 And A_bout(boutCount) >= B_bout(boutCount) + 2 Then
                        A_set = A_set + 1
                        A_game(gameCount) = A_game(gameCount) + 1
                        gameCount = gameCount + 1
                        If gameCount > UBound(A_game) Then
                              ReDim Preserve A_game(gameCount + 5)
                              ReDim Preserve B_game(gameCount + 5)
                        End If
                        boutCount = boutCount + 1
                        If boutCount > UBound(A_bout) Then
                              ReDim Preserve A_bout(boutCount + 50)
                              ReDim Preserve B_bout(boutCount + 5)
                        End If
                        If strExam(exchangeStr) = "1" Then
                              exchangeFlag = exchangeFlag * (-1)
                        End If
                  ElseIf B_bout(boutCount) >= 7 And B_bout(boutCount) >= A_bout(boutCount) + 2 Then
                        B_set = B_set + 1
                        B_game(gameCount) = B_game(gameCount) + 1
                        gameCount = gameCount + 1
                        If gameCount > UBound(A_game) Then
                              ReDim Preserve A_game(gameCount + 5)
                              ReDim Preserve B_game(gameCount + 5)
                        End If
                        boutCount = boutCount + 1
                        If boutCount > UBound(A_bout) Then
                              ReDim Preserve A_bout(boutCount + 50)
                              ReDim Preserve B_bout(boutCount + 5)
                        End If
                        If strExam(exchangeStr) = "1" Then
                              exchangeFlag = exchangeFlag * (-1)
                        End If
                  End If
            Else
普通局:
                  If A_bout(boutCount) >= 4 And A_bout(boutCount) >= B_bout(boutCount) + 2 Then
                        A_game(gameCount) = A_game(gameCount) + 1
                        boutCount = boutCount + 1
                        If boutCount > UBound(A_bout) Then
                              ReDim Preserve A_bout(boutCount + 50)
                              ReDim Preserve B_bout(boutCount + 5)
                        End If
                        If strExam(exchangeStr) = "1" Then
                              exchangeFlag = exchangeFlag * (-1)
                        End If
                        If A_game(gameCount) >= gameMode And A_game(gameCount) >= B_game(gameCount) + 2 Then
                              A_set = A_set + 1
                              gameCount = gameCount + 1
                              If gameCount > UBound(A_game) Then
                                    ReDim Preserve A_game(gameCount + 5)
                                    ReDim Preserve B_game(gameCount + 5)
                              End If
                        End If
                  ElseIf B_bout(boutCount) >= 4 And B_bout(boutCount) >= A_bout(boutCount) + 2 Then
                        B_game(gameCount) = B_game(gameCount) + 1
                        boutCount = boutCount + 1
                        If boutCount > UBound(A_bout) Then
                              ReDim Preserve A_bout(boutCount + 50)
                              ReDim Preserve B_bout(boutCount + 5)
                        End If
                        If strExam(exchangeStr) = "1" Then
                              exchangeFlag = exchangeFlag * (-1)
                        End If
                        If B_game(gameCount) >= gameMode And B_game(gameCount) >= A_game(gameCount) + 2 Then
                              B_set = B_set + 1
                              gameCount = gameCount + 1
                              If gameCount > UBound(A_game) Then
                                    ReDim Preserve A_game(gameCount + 5)
                                    ReDim Preserve B_game(gameCount + 5)
                              End If
                        End If
                  End If
            End If
      Else
      End If
'''''''找破发及破发点用
'      Debug.Print boutCount & " : " & A_bout(boutCount) & ":" & B_bout(boutCount)
End Sub

Private Function strExam(str As String) As String
      If str = "" Then
            strExam = "0"
      Else
            strExam = Mid(str, 1, 1)
            str = Mid(str, 2, Len(str) - 1)
      End If
End Function

Public Sub showBout(rng As Range)
      Dim i
      rng.Value = "回合分A:B"
      For i = 0 To boutCount - 1 Step 1
            With rng.Offset(i + 1)
                  .NumberFormatLocal = "@"
                  .Value = A_bout(i) & " : " & B_bout(i)
            End With
      Next
End Sub

Public Sub showGame(rng As Range)
      Dim i
      rng.Value = "局分A:B"
      For i = 0 To gameCount - 1 Step 1
            With rng.Offset(i + 1)
                  .NumberFormatLocal = "@"
                  .Value = A_game(i) & " : " & B_game(i)
            End With
      Next
End Sub

Public Sub showSet(rng As Range)
      rng.Value = "盘分A:B"
      rng.Offset(1).NumberFormatLocal = "@"
      rng.Offset(1).Value = A_set & " : " & B_set
End Sub

