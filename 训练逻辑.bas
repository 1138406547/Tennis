Attribute VB_Name = "ģ��1"
Option Explicit

Const WAITING As Integer = 19
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''ѵ���߼�����'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub JudgeTrainning(current As ball, last As ball, useful As ball, i%, trainningMode%, Optional trainningSide% = 1)
      
''''''��ѵ���߼��У���ѵ��˫���滮Ϊѵ���ߺ������ߣ�
''''''ѵ��ģʽ(trainningMode)�Լ����ڷ�λ(trainningSide)
''''''ȷ�����������ݼ�¼��A������B�������¡�

      Dim ServeTargetAreaEasyAntiA1 As New multiZone                              '��A1������ķ���ѵ��-��Ŀ������
      Dim ServeTargetAreaEasyAntiA2 As New multiZone                              '��A2������ķ���ѵ��-��Ŀ������
      Dim ServeTargetAreaEasyAntiB1 As New multiZone                              '��B1������ķ���ѵ��-��Ŀ������
      Dim ServeTargetAreaEasyAntiB2 As New multiZone                              '��B2������ķ���ѵ��-��Ŀ������
      
      Dim HitTargetAreaEasyAntiA As New multiZone                                 '��A���ػ����Ѷ���Ŀ������
      Dim HitTargetAreaEasyAntiB As New multiZone                                 '��B���ػ����Ѷ���Ŀ������
      

      Dim tmpZone As New zone
      Dim tempMultiZone As New multiZone
      
      If Sheets("main").Range("N3").Value = 1 Then                      'cm����
            Call tmpZone.init(0, 411, 640, 0)
            Call ServeTargetAreaEasyAntiB1.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call ServeTargetAreaEasyAntiB2.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call ServeTargetAreaEasyAntiA1.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call ServeTargetAreaEasyAntiA2.push_back(tmpZone)
            
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 411, 1188, -411)
            Call HitTargetAreaEasyAntiB.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call HitTargetAreaEasyAntiA.push_back(tmpZone)

      ElseIf Sheets("main").Range("N3").Value = 2 Then                  'mm����
            Call tmpZone.init(0, 4115, 6401, 0)
            Call ServeTargetAreaEasyAntiB1.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call ServeTargetAreaEasyAntiB2.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call ServeTargetAreaEasyAntiA1.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call ServeTargetAreaEasyAntiA2.push_back(tmpZone)
            
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 4115, 11887, -4115)
            Call HitTargetAreaEasyAntiB.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call HitTargetAreaEasyAntiA.push_back(tmpZone)
      End If
      
      Dim commit$, col%       '�������Լ��������������б�
      If trainningSide > 0 Then
            col = 6
      Else
            col = 9
      End If
      commit = "waitingForNextServe"
      Dim ba As New ball
      
      
      With Sheets("main")
            
            Select Case current.stat
            Case 1
                  Select Case trainningMode
                  Case 1      '����ѵ��
                        last.stat = WAITING
                        commit = "NextServe"
                        Call description(.Cells(i, col), commit)
                  Case 2      '�ӷ���ѵ��
                        If last.stat = 3 Then
                              If current.x * trainningSide > 0 Then
                                    Call last.clone(current)
                                    commit = "Return"
                                    Call description(.Cells(i, col), commit)
                              Else
                                    last.stat = WAITING
                                    commit = "11111111111111111111111111111111111111111111111111111111"
                                    Call description(.Cells(i, col), commit)
                              End If
                        Else
                              last.stat = WAITING
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case 3      '����ѵ��
                        If current.x * trainningSide < 0 Then   '��ǰΪ�Է�����
                              Call last.clone(current)
                              commit = "Feed"
                              Call description(.Cells(i, col), commit)
                        Else                                      '��ǰΪ��������
                              Select Case last.stat
                              Case 3
                                    Call last.clone(current)
                                    commit = "Hit"
                                    Call description(.Cells(i, col), commit)
                              Case Else
                                    last.stat = WAITING
                                    commit = "NextFeed"
                                    Call description(.Cells(i, col), commit)
                              End Select
                        End If
                  Case Else
                        MsgBox "error"
                        Exit Sub
                  End Select
                        
            
            Case 2
                  Select Case trainningMode
                  Case 1      '����ѵ��
                        If current.x * trainningSide > 0 And last.stat = WAITING Then
                              Call last.clone(current)
                              commit = "Serve"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case 2      '�ӷ���
                        If current.x * trainningSide < 0 And last.stat = WAITING Then
                              Call last.clone(current)
                              commit = "Serve"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case 3      '����
                        If current.x * trainningSide < 0 Then
                              Call last.clone(current)
                              last.stat = 1
                              commit = "Feed"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextFeed"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case Else
                        MsgBox "error"
                        Exit Sub
                  End Select
                        
                        
            Case 3
                  Select Case trainningMode
                  Case 1      '����ѵ��
                        If last.stat = 2 Then
                              If (last.x > 0 And last.y >= 0 And current.isInside(ServeTargetAreaEasyAntiA1)) Or _
                                 (last.x > 0 And last.y < 0 And current.isInside(ServeTargetAreaEasyAntiA2)) Or _
                                 (last.x < 0 And last.y >= 0 And current.isInside(ServeTargetAreaEasyAntiB2)) Or _
                                 (last.x < 0 And last.y < 0 And current.isInside(ServeTargetAreaEasyAntiB1)) Then
                                    commit = "ServeIn"
                                    Call description(.Cells(i, col), commit)
                              Else
                                    commit = "Fault"
                                    Call description(.Cells(i, col), commit)
                              End If
                        ElseIf last.stat = 4 And useful.stat = 2 Then
                              If (useful.x > 0 And useful.y >= 0 And current.isInside(ServeTargetAreaEasyAntiA1)) Or _
                                 (useful.x > 0 And useful.y < 0 And current.isInside(ServeTargetAreaEasyAntiA2)) Or _
                                 (useful.x < 0 And useful.y >= 0 And current.isInside(ServeTargetAreaEasyAntiB2)) Or _
                                 (useful.x < 0 And useful.y < 0 And current.isInside(ServeTargetAreaEasyAntiB1)) Then
                                    commit = "Let"
                                    Call description(.Cells(i, col), commit)
                              Else
                                    commit = "Fault"
                                    Call description(.Cells(i, col), commit)
                              End If
                        ElseIf last.stat = 9 And useful.stat = 2 Then
                              If (useful.x > 0 And useful.y >= 0 And current.isInside(ServeTargetAreaEasyAntiA1)) Or _
                                 (useful.x > 0 And useful.y < 0 And current.isInside(ServeTargetAreaEasyAntiA2)) Or _
                                 (useful.x < 0 And useful.y >= 0 And current.isInside(ServeTargetAreaEasyAntiB2)) Or _
                                 (useful.x < 0 And useful.y < 0 And current.isInside(ServeTargetAreaEasyAntiB1)) Then
                                    commit = "ServeIn"
                                    Call description(.Cells(i, col), commit)
                              Else
                                    commit = "Fault"
                                    Call description(.Cells(i, col), commit)
                              End If
                        Else
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                        last.stat = WAITING
                  Case 2            '�ӷ���ѵ��
                        If last.stat = 1 Then  'ѵ���߻���
                              If (last.x > 0 And current.isInside(HitTargetAreaEasyAntiA)) Or _
                                 (last.x < 0 And current.isInside(HitTargetAreaEasyAntiB)) Then
                                    commit = "In"
                                    Call description(.Cells(i, col), commit)
                              Else
                                    commit = "Out"
                                    Call description(.Cells(i, col), commit)
                              End If
                              last.stat = WAITING
                        ElseIf last.stat = 2 Then
                              If (last.x > 0 And last.y >= 0 And current.isInside(ServeTargetAreaEasyAntiA1)) Or _
                                 (last.x > 0 And last.y < 0 And current.isInside(ServeTargetAreaEasyAntiA2)) Or _
                                 (last.x < 0 And last.y >= 0 And current.isInside(ServeTargetAreaEasyAntiB2)) Or _
                                 (last.x < 0 And last.y < 0 And current.isInside(ServeTargetAreaEasyAntiB1)) Then
                                    Call last.clone(current)
                                    commit = "ServeIn"
                                    Call description(.Cells(i, col), commit)
                              Else
                                    last.stat = WAITING
                                    commit = "Fault"
                                    Call description(.Cells(i, col), commit)
                              End If
                        ElseIf last.stat = 4 Or last.stat = 9 Then
                              Call ba.clone(useful)
                              Call JudgeTrainning(current, ba, useful, i, trainningMode, trainningSide)
                              Call last.clone(ba)
                        Else
                              last.stat = WAITING
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case 3            '����ѵ��
                        If last.stat = 1 Then
                              If last.x * trainningSide > 0 Then        '��һ������Ϊѵ���߻���
                                    If (last.x > 0 And current.isInside(HitTargetAreaEasyAntiA)) Or _
                                       (last.x < 0 And current.isInside(HitTargetAreaEasyAntiB)) Then
                                          commit = "In"
                                          Call description(.Cells(i, col), commit)
                                    Else
                                          commit = "Out"
                                          Call description(.Cells(i, col), commit)
                                    End If
                                    last.stat = WAITING
                              Else        '��һ������Ϊι���߻���
                                    If (last.x > 0 And current.isInside(HitTargetAreaEasyAntiA)) Or _
                                       (last.x < 0 And current.isInside(HitTargetAreaEasyAntiB)) Then
                                          Call last.clone(current)
                                          commit = "FeedIn"
                                          Call description(.Cells(i, col), commit)
                                    Else
                                          last.stat = WAITING
                                          commit = "FeedOut"
                                          Call description(.Cells(i, col), commit)
                                    End If
                              End If
                        ElseIf last.stat = 4 Or last.stat = 9 Then
                              Call ba.clone(useful)
                              Call JudgeTrainning(current, ba, useful, i, trainningMode, trainningSide)
                              Call last.clone(ba)
                        Else
                              last.stat = WAITING
                              commit = "NextFeed"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case Else
                        MsgBox "error"
                        Exit Sub
                  End Select
            
            
            Case 4
                  Select Case trainningMode
                  Case 1      '����ѵ��
                        If last.stat = 2 Then
                              Call useful.clone(last)
                              Call last.clone(current)
                              commit = "net"
                              Call description(.Cells(i, col), commit)
                        ElseIf last.stat = 9 Then
                              Call last.clone(current)
                              commit = "net"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case 2      '�ӷ���ѵ��
                        If last.stat = 1 Then
                              Call useful.clone(last)
                              Call last.clone(current)
                              commit = "net"
                              Call description(.Cells(i, col), commit)
                        ElseIf last.stat = 9 Then
                              Call last.clone(current)
                              commit = "net"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case 3      '����ѵ��
                        If last.stat = 1 Then
                              Call useful.clone(last)
                              Call last.clone(current)
                              commit = "net"
                              Call description(.Cells(i, col), commit)
                        ElseIf last.stat = 9 Then
                              Call last.clone(current)
                              commit = "net"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextFeed"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case Else
                        MsgBox "error"
                        Exit Sub
                  End Select
                  
                  
            Case 6
                  Select Case trainningMode
                  Case 1      '����ѵ��
                        If last.stat = 2 Then
                              Call useful.clone(last)
                              Call last.clone(current)
                              commit = "OutOfSight"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case 2      '�ӷ���ѵ��
                        If last.stat = 1 Or last.stat = 2 Or last.stat = 3 Then
                              Call useful.clone(last)
                              Call last.clone(current)
                              commit = "OutOfSight"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextServe"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case 3      '����ѵ��
                        If last.stat = 1 Or last.stat = 3 Then
                              Call useful.clone(last)
                              Call last.clone(current)
                              commit = "OutOfSight"
                              Call description(.Cells(i, col), commit)
                        Else
                              last.stat = WAITING
                              commit = "NextFeed"
                              Call description(.Cells(i, col), commit)
                        End If
                  Case Else
                        MsgBox "error"
                        Exit Sub
                  End Select
                              
                              
            Case 9
                  If last.stat = 6 Then
                        Call last.clone(current)
                        commit = "BackInSight"
                        Call description(.Cells(i, col), commit)
                  Else
                        last.stat = WAITING
                        If trainningMode = 3 Then
                              commit = "NextFeed"
                        Else
                              commit = "NextServe"
                        End If
                        Call description(.Cells(i, col), commit)
                  End If
                  
                  
            Case 10
                  If trainningMode = 1 Then
                        last.stat = WAITING
                        commit = "NextServe"
                        Call description(.Cells(i, col), commit)
                  Else
                        current.stat = 1
                        Call ba.clone(useful)
                        Call JudgeTrainning(current, ba, useful, i, trainningMode, trainningSide)
                        Call last.clone(ba)
                  End If
                  
            End Select
      End With
End Sub

Private Sub description(rng As Range, words As String)
      rng.Value = words
End Sub



