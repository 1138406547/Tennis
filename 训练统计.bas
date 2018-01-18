Attribute VB_Name = "ģ��2"
Option Explicit

Const WAITING As Integer = 19
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''ѵ��ͳ�Ʋ���'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub TrainningAnalysis(aimColumn%, trainningMode%, Optional trainningSide% = 1, Optional trainningLevel = 1)
      
''''''��ѵ���߼��У���ѵ��˫���滮Ϊѵ���ߺ������ߣ�
''''''ѵ��ģʽ(trainningMode)�Լ����ڷ�λ(trainningSide)
''''''ȷ�����������ݼ�¼��A������B�������¡�

      
      Dim ServeInEasy As New vecBall                                              '��ǰѵ���߷���ѵ��-�׽���������
      Dim ServeInNormal As New vecBall                                            '��ǰѵ���߷���ѵ��-�н���������
      Dim ServeInHard As New vecBall                                              '��ǰѵ���߷���ѵ��-�ѽ���������
      Dim ServeTargetEasy As New vecBall                                          '��ǰѵ���߷���ѵ��-����Ŀ��������
      Dim ServeTargetNormal As New vecBall                                        '��ǰѵ���߷���ѵ��-����Ŀ��������
      Dim ServeTargetHard As New vecBall                                          '��ǰѵ���߷���ѵ��-����Ŀ��������
      Dim ServePoint As New vecBall                                               '��ǰѵ���߷���ѵ����ѵ������
      Dim ServeLandingPointFault As New vecBall                                   '��ǰѵ���߷���ѵ��ʧ�ܴ���
      Dim ServeLetPoint As New vecBall                                            '��ǰѵ���߷���ѵ������
      
      Dim ReturnInEasy As New vecBall                                             '��ǰѵ���߽ӷ���ѵ��-�׽���������
      Dim ReturnInNormal As New vecBall                                           '��ǰѵ���߽ӷ���ѵ��-�н���������
      Dim ReturnInHard As New vecBall                                             '��ǰѵ���߽ӷ���ѵ��-�ѽ���������
      Dim ReturnTargetEasy As New vecBall                                         '��ǰѵ���߽ӷ���ѵ��-����Ŀ��������
      Dim ReturnTargetNormal As New vecBall                                       '��ǰѵ���߽ӷ���ѵ��-����Ŀ��������
      Dim ReturnTargetHard As New vecBall                                         '��ǰѵ���߽ӷ���ѵ��-����Ŀ��������
      Dim ReturnPoint As New vecBall                                              '��ǰѵ���߽ӷ���ѵ���ӷ���������
      Dim ReturnOutPoint As New vecBall                                           '��ǰѵ���߽ӷ���ѵ���������
      Dim ReturnTotalPoint As New vecBall                                         '��ǰѵ���߽ӷ���ѵ����ѵ������
      
      Dim HitInEasy As New vecBall                                                '��ǰѵ���߻���ѵ��-�׽���������
      Dim HitInNormal As New vecBall                                              '��ǰѵ���߻���ѵ��-�н���������
      Dim HitInHard As New vecBall                                                '��ǰѵ���߻���ѵ��-�ѽ���������
      Dim HitTargetEasy As New vecBall                                            '��ǰѵ���߻���ѵ��-����Ŀ��������
      Dim HitTargetNormal As New vecBall                                          '��ǰѵ���߻���ѵ��-����Ŀ��������
      Dim HitTargetHard As New vecBall                                            '��ǰѵ���߻���ѵ��-����Ŀ��������
      Dim HitOutPoint As New vecBall                                              '��ǰѵ���߻���ѵ���������
      Dim HitPoint As New vecBall                                                 '��ǰѵ���߻���ѵ����ѵ������
      Dim HitBeingVolleyPoint As New vecBall                                      '��ǰѵ���߻��򱻶Է��ػ�
      
      Dim ServeTargetAreaNormal As New multiZone                                  '����ѵ��-��Ŀ������
      Dim ServeTargetAreaHard As New multiZone                                    '����ѵ��-��Ŀ������
      Dim inner As New multiZone                                  '����_�ڽ�����
      Dim medium As New multiZone                                 '����_��·����
      Dim outer As New multiZone                                  '����_�������
      Dim innerFaultAntiB1 As New multiZone                       '��B1������ʱ�ж�Ϊ����_�ڽ�ʧ�������
      Dim innerFaultAntiB2 As New multiZone                       '��B2������ʱ�ж�Ϊ����_�ڽ�ʧ�������
      Dim innerFaultAntiA1 As New multiZone                       '��A1������ʱ�ж�Ϊ����_�ڽ�ʧ�������
      Dim innerFaultAntiA2 As New multiZone                       '��A2������ʱ�ж�Ϊ����_�ڽ�ʧ�������
      Dim mediumFaultAntiB1 As New multiZone                      '��B1������ʱ�ж�Ϊ����_��·ʧ�������
      Dim mediumFaultAntiB2 As New multiZone                      '��B2������ʱ�ж�Ϊ����_��·ʧ�������
      Dim mediumFaultAntiA1 As New multiZone                      '��A1������ʱ�ж�Ϊ����_��·ʧ�������
      Dim mediumFaultAntiA2 As New multiZone                      '��A2������ʱ�ж�Ϊ����_��·ʧ�������
      Dim outerFaultAntiB1 As New multiZone                       '��B1������ʱ�ж�Ϊ����_���ʧ�������
      Dim outerFaultAntiB2 As New multiZone                       '��B2������ʱ�ж�Ϊ����_���ʧ�������
      Dim outerFaultAntiA1 As New multiZone                       '��A1������ʱ�ж�Ϊ����_���ʧ�������
      Dim outerFaultAntiA2 As New multiZone                       '��A2������ʱ�ж�Ϊ����_���ʧ�������

      
      Dim ReturnTargetAreaEasyAntiA As New multiZone                              '������A������ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaNormalAntiA1 As New multiZone                           '������A1����ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaHardAntiA1 As New multiZone                             '������A1����ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaNormalAntiA2 As New multiZone                           '������A2����ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaHardAntiA2 As New multiZone                             '������A2����ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaEasyAntiB As New multiZone                              '������B������ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaNormalAntiB1 As New multiZone                           '������B1����ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaHardAntiB1 As New multiZone                             '������B1����ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaNormalAntiB2 As New multiZone                           '������B2����ʱ�ӷ���ѵ����Ŀ������
      Dim ReturnTargetAreaHardAntiB2 As New multiZone                             '������B2����ʱ�ӷ���ѵ����Ŀ������
      
      Dim HitTargetAreaEasyAntiA As New multiZone                                 '��A���ػ����Ѷ���Ŀ������
      Dim HitTargetAreaEasyAntiB As New multiZone                                 '��B���ػ����Ѷ���Ŀ������
      Dim HitTargetAreaNormal As New multiZone                                    '�����Ѷ���Ŀ������
      Dim HitTargetAreaHard As New multiZone                                      '�����Ѷ���Ŀ������
      
      Dim tmpZone As New zone
      Dim tempMultiZone As New multiZone
      Dim FOREBACK As Integer                   '�����ַֽ��߾���ֵ
      If Sheets("main").Range("N3").Value = 1 Then                      'cm����
            FOREBACK = 274
            Call tmpZone.init(0, 50, 640, 0)                      '����_�ڽ���A1���ķ�Χ
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorX                                  '����_�ڽ���A2���ķ�Χ
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorY                                  '����_�ڽ���B1���ķ�Χ
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorX                                  '����_�ڽ���B2���ķ�Χ
            Call inner.push_back(tmpZone)
            Call tmpZone.init(0, 361, 640, 50)
            Call medium.push_back(tmpZone)                        '��������A1���ķ�Χ
            Call tmpZone.mirrorX
            Call medium.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call medium.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call medium.push_back(tmpZone)
            Call tmpZone.init(0, 411, 640, 361)
            Call outer.push_back(tmpZone)                         '��������A1���ķ�Χ
            Call tmpZone.mirrorX
            Call outer.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call outer.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call outer.push_back(tmpZone)
            
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 0, 1588, -50)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(640, 50, 1588, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call innerFaultAntiB1.clone(tempMultiZone)            '��B1������ʱ�ж�Ϊ����_�ڽ�ʧ�������
            Call tempMultiZone.mirrorX
            Call innerFaultAntiB2.clone(tempMultiZone)            '��B2������ʱ�ж�Ϊ����_�ڽ�ʧ�������
            Call tempMultiZone.mirrorY
            Call innerFaultAntiA1.clone(tempMultiZone)            '��A1������ʱ�ж�Ϊ����_�ڽ�ʧ�������
            Call tempMultiZone.mirrorX
            Call innerFaultAntiA2.clone(tempMultiZone)            '��A2������ʱ�ж�Ϊ����_�ڽ�ʧ�������
            Call tmpZone.init(640, 361, 1588, 50)
            Call mediumFaultAntiB1.push_back(tmpZone)             '��B1������ʱ�ж�Ϊ����_��·ʧ�������
            Call tmpZone.mirrorX
            Call mediumFaultAntiB2.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call mediumFaultAntiA1.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call mediumFaultAntiA2.push_back(tmpZone)
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 461, 1588, 411)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(640, 411, 1588, 361)
            Call tempMultiZone.push_back(tmpZone)
            Call outerFaultAntiB1.clone(tempMultiZone)            '��B1������ʱ�ж�Ϊ����_���ʧ�������
            Call tempMultiZone.mirrorX
            Call outerFaultAntiB2.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call outerFaultAntiA1.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call outerFaultAntiA2.combine(tempMultiZone)
            
            Call tmpZone.init(0, 411, 640, 361)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(457, 361, 640, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(0, 50, 640, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call ServeTargetAreaHard.clone(tempMultiZone)               '������Ŀ������
            Call tempMultiZone.mirrorX
            Call ServeTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call ServeTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call ServeTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 411, 640, 361)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(240, 361, 640, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(0, 50, 640, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call ServeTargetAreaNormal.clone(tempMultiZone)             '������Ŀ������
            Call tempMultiZone.mirrorX
            Call ServeTargetAreaNormal.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call ServeTargetAreaNormal.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call ServeTargetAreaNormal.combine(tempMultiZone)
            
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 411, 1188, -411)
            Call ReturnTargetAreaEasyAntiB.push_back(tmpZone)           '��B���ӷ���ʱ�Ѷ�Ϊ�׵�Ŀ������
            Call HitTargetAreaEasyAntiB.push_back(tmpZone)              '��B����ʱ�Ѷ�Ϊ�׵�Ŀ������
            Call tmpZone.mirrorY
            Call ReturnTargetAreaEasyAntiA.push_back(tmpZone)           '��A���ӷ���ʱ�Ѷ�Ϊ�׵�Ŀ������
            Call HitTargetAreaEasyAntiA.push_back(tmpZone)              '��A����ʱ�Ѷ�Ϊ�׵�Ŀ������
            Call tempMultiZone.Clear
            Call tmpZone.init(-1188, 411, 0, 229)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-1188, 411, -1006, -411)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-1188, -361, 0, -411)
            Call tempMultiZone.push_back(tmpZone)
            Call ReturnTargetAreaHardAntiA1.clone(tempMultiZone)        '��A1���ӷ���ʱ�Ѷ�Ϊ�ѵ�Ŀ������
            Call tempMultiZone.mirrorX
            Call ReturnTargetAreaHardAntiA2.clone(tempMultiZone)        '��A2���ӷ���ʱ�Ѷ�Ϊ�ѵ�Ŀ������
            Call tempMultiZone.mirrorY
            Call ReturnTargetAreaHardAntiB1.clone(tempMultiZone)        '��B1���ӷ���ʱ�Ѷ�Ϊ�ѵ�Ŀ������
            Call tempMultiZone.mirrorX
            Call ReturnTargetAreaHardAntiB2.clone(tempMultiZone)        '��B2���ӷ���ʱ�Ѷ�Ϊ�ѵ�Ŀ������
            Call tempMultiZone.Clear
            Call tmpZone.init(-1188, 411, 0, 229)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-1188, 411, -789, -411)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-1188, -361, 0, -411)
            Call tempMultiZone.push_back(tmpZone)
            Call ReturnTargetAreaNormalAntiA1.clone(tempMultiZone)      '��A1���ӷ���ʱ�Ѷ�Ϊ�е�Ŀ������
            Call tempMultiZone.mirrorX
            Call ReturnTargetAreaNormalAntiA2.clone(tempMultiZone)      '��A2���ӷ���ʱ�Ѷ�Ϊ�е�Ŀ������
            Call tempMultiZone.mirrorY
            Call ReturnTargetAreaNormalAntiB1.clone(tempMultiZone)      '��B1���ӷ���ʱ�Ѷ�Ϊ�е�Ŀ������
            Call tempMultiZone.mirrorX
            Call ReturnTargetAreaNormalAntiB2.clone(tempMultiZone)      '��B2���ӷ���ʱ�Ѷ�Ϊ�е�Ŀ������
            
            Call tempMultiZone.Clear
            Call tmpZone.init(-1188, 411, 0, 361)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-1188, 411, -1006, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call HitTargetAreaHard.clone(tempMultiZone)                 '�����Ѷ�Ϊ�ѵ�Ŀ������
            Call tempMultiZone.mirrorX
            Call HitTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call HitTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call HitTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.Clear
            Call tmpZone.init(-1188, 411, 0, 361)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-1188, 411, -789, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call HitTargetAreaNormal.clone(tempMultiZone)               '�����Ѷ�Ϊ�е�Ŀ������
            Call tempMultiZone.mirrorX
            Call HitTargetAreaNormal.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call HitTargetAreaNormal.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call HitTargetAreaNormal.combine(tempMultiZone)
      ElseIf Sheets("main").Range("N3").Value = 2 Then                      'mm����
            FOREBACK = 2744
            Call tmpZone.init(0, 500, 6401, 0)
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call inner.push_back(tmpZone)
            Call tmpZone.init(0, 3615, 6401, 500)
            Call medium.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call medium.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call medium.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call medium.push_back(tmpZone)
            Call tmpZone.init(0, 4115, 6401, 3615)
            Call outer.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call outer.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call outer.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call outer.push_back(tmpZone)
            
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 0, 15887, -500)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(6401, 500, 15887, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call innerFaultAntiB1.clone(tempMultiZone)            '��B1������ʱ�ж�Ϊ����_�ڽ�ʧ�������
            Call tempMultiZone.mirrorX
            Call innerFaultAntiB2.clone(tempMultiZone)            '��B2������ʱ�ж�Ϊ����_�ڽ�ʧ�������
            Call tempMultiZone.mirrorY
            Call innerFaultAntiA1.clone(tempMultiZone)            '��A1������ʱ�ж�Ϊ����_�ڽ�ʧ�������
            Call tempMultiZone.mirrorX
            Call innerFaultAntiA2.clone(tempMultiZone)            '��A2������ʱ�ж�Ϊ����_�ڽ�ʧ�������
            Call tmpZone.init(6401, 3615, 15887, 500)
            Call mediumFaultAntiB1.push_back(tmpZone)             '��B1������ʱ�ж�Ϊ����_��·ʧ�������
            Call tmpZone.mirrorX
            Call mediumFaultAntiB2.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call mediumFaultAntiA1.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call mediumFaultAntiA2.push_back(tmpZone)
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 4615, 15887, 4115)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(6401, 4115, 15887, 3615)
            Call tempMultiZone.push_back(tmpZone)
            Call outerFaultAntiB1.clone(tempMultiZone)            '��B1������ʱ�ж�Ϊ����_���ʧ�������
            Call tempMultiZone.mirrorX
            Call outerFaultAntiB2.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call outerFaultAntiA1.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call outerFaultAntiA2.combine(tempMultiZone)
            
            Call tmpZone.init(0, 4115, 6401, 3615)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(4571, 3615, 6401, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(0, 500, 6401, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call ServeTargetAreaHard.clone(tempMultiZone)                     '������Ŀ������
            Call tempMultiZone.mirrorX
            Call ServeTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call ServeTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call ServeTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 4115, 6401, 3615)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(2400, 3615, 6401, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(0, 500, 6401, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call ServeTargetAreaNormal.clone(tempMultiZone)                   '������Ŀ������
            Call tempMultiZone.mirrorX
            Call ServeTargetAreaNormal.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call ServeTargetAreaNormal.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call ServeTargetAreaNormal.combine(tempMultiZone)
            
            Call tempMultiZone.Clear
            Call tmpZone.init(0, 4115, 11887, -4115)
            Call ReturnTargetAreaEasyAntiB.push_back(tmpZone)
            Call HitTargetAreaEasyAntiB.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call ReturnTargetAreaEasyAntiA.push_back(tmpZone)
            Call HitTargetAreaEasyAntiA.push_back(tmpZone)
            Call tempMultiZone.Clear
            Call tmpZone.init(-11887, 4115, 0, 2286)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-11887, 4115, -10059, -4115)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-11887, -3621, 0, -4115)
            Call tempMultiZone.push_back(tmpZone)
            Call ReturnTargetAreaHardAntiA1.clone(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call ReturnTargetAreaHardAntiA2.clone(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call ReturnTargetAreaHardAntiB1.clone(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call ReturnTargetAreaHardAntiB2.clone(tempMultiZone)
            Call tempMultiZone.Clear
            Call tmpZone.init(-11887, 4115, 0, 2286)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-11887, 4115, -7890, -4115)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-11887, -3621, 0, -4115)
            Call tempMultiZone.push_back(tmpZone)
            Call ReturnTargetAreaNormalAntiA1.clone(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call ReturnTargetAreaNormalAntiA2.clone(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call ReturnTargetAreaNormalAntiB1.clone(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call ReturnTargetAreaNormalAntiB2.clone(tempMultiZone)
            
            Call tempMultiZone.Clear
            Call tmpZone.init(-11887, 4115, 0, 3615)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-11887, 4115, -10059, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call HitTargetAreaHard.clone(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call HitTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call HitTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call HitTargetAreaHard.combine(tempMultiZone)
            Call tempMultiZone.Clear
            Call tmpZone.init(-11887, 4115, 0, 3615)
            Call tempMultiZone.push_back(tmpZone)
            Call tmpZone.init(-11887, 4115, -7890, 0)
            Call tempMultiZone.push_back(tmpZone)
            Call HitTargetAreaNormal.clone(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call HitTargetAreaNormal.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call HitTargetAreaNormal.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call HitTargetAreaNormal.combine(tempMultiZone)
      End If
      
      Dim ReturnForehandTotal As New vecBall                '�ӷ������ֻ����
      Dim ReturnBackhandTotal As New vecBall                '�ӷ����ֻ����
      Dim HitForehandTotal As New vecBall                   '�������ֻ����
      Dim HitBackhandTotal As New vecBall                   '�����ֻ����
      
      Dim ReturnForehandIn As New vecBall                   '�ӷ��������������
      Dim ReturnBackhandIn As New vecBall                   '�ӷ������������
      Dim HitForehandIn As New vecBall                      '���������������
      Dim HitBackhandIn As New vecBall                      '�������������
      
      Dim ServeLandingOtherFault As New vecBall             '����ʧ���������
      Dim ServeLandingInnerFault As New vecBall             '�����ڽ�ʧ�����
      Dim ServeLandingMediumFault As New vecBall            '������·ʧ�����
      Dim ServeLandingOuterFault As New vecBall             '�������ʧ�����
      Dim ServeLandingInner As New vecBall                  '�������_��
      Dim ServeLandingMedium As New vecBall                 '�������_��
      Dim ServeLandingOuter As New vecBall                  '�������_��
      
      
      
      Dim commit$             '�������Լ��������������б�
      
      Dim ba As New ball
      Dim aim As Range
      Dim i%, j%, k%
      Dim latestHit As New ball                                     '��һ����������(����1������2)
      Dim maxSucceed%, tempSucceed%                               '��������ɹ�
      Dim shortRound%, middleRound%, longRound%, roundCount%      '����,�г���,����
      Dim islatestHitTouchDown As Boolean                           '�����ж�����ѵ�����Ƿ񱻽ػ�
      Dim islatestHitForehand As Boolean                            '��һ����������������ʣ�����true;���ֻ���false
      islatestHitTouchDown = True
      maxSucceed = 0
      tempSucceed = 0
      
      With Sheets("main")
            Set aim = .Range(.Cells(1, aimColumn), .Cells(9999, aimColumn).End(xlUp))
            For i = 1 To aim.Cells.count
                  If aim.Cells(i) <> "" Then
                        j = aim.Cells(i).Row
                        Call ba.init(.Cells(j, 1), .Cells(j, 2), .Cells(j, 3), .Cells(j, 4), .Cells(j, 5))
                        If trainningMode = 1 Then     '����ѵ��----------------------------------------------------------------------------------
                              If .Cells(j, aimColumn) = "Serve" Then          '����------------------
                                    Call ServePoint.push_back(ba)
                                    Call latestHit.clone(ba)
                              ElseIf .Cells(j, aimColumn) = "ServeIn" Then    '��������----------------
'''������������Ŀ�������ж�
                                    Call ServeTargetEasy.push_back(ba)
                                    Call ServeInEasy.push_back(ba)
                                    commit = "Easy"
                                    If trainningLevel = 1 Then
                                          tempSucceed = tempSucceed + 1
                                    End If
                                    If ba.isInside(ServeTargetAreaNormal) Then
                                          Call ServeTargetNormal.push_back(ba)
                                          commit = "Normal"
                                          If trainningLevel = 2 Then
                                                tempSucceed = tempSucceed + 1
                                          End If
                                    Else
                                          Call ServeInNormal.push_back(ba)
                                          If trainningLevel = 2 Then
                                                If tempSucceed > maxSucceed Then
                                                      maxSucceed = tempSucceed
                                                End If
                                                tempSucceed = 0
                                          End If
                                    End If
                                    If ba.isInside(ServeTargetAreaHard) Then
                                          Call ServeTargetHard.push_back(ba)
                                          commit = "Hard"
                                          If trainningLevel = 3 Then
                                                tempSucceed = tempSucceed + 1
                                          End If
                                    Else
                                          If trainningLevel = 3 Then
                                                If tempSucceed > maxSucceed Then
                                                      maxSucceed = tempSucceed
                                                End If
                                                tempSucceed = 0
                                          End If
                                          Call ServeInHard.push_back(ba)
                                    End If
'''
                                    Call description(.Cells(j, aimColumn + 1), commit)
'''��������������ж�
                                    If ba.isInside(inner) Then
                                          Call ServeLandingInner.push_back(latestHit)
                                    ElseIf ba.isInside(medium) Then
                                          Call ServeLandingMedium.push_back(latestHit)
                                    ElseIf ba.isInside(outer) Then
                                          Call ServeLandingOuter.push_back(latestHit)
                                    Else
                                          commit = "111111111111111111111111111111111111111111111"
                                          Call description(.Cells(i, aimColumn + 1), commit)
                                    End If
'''
                              ElseIf .Cells(j, aimColumn) = "Fault" Then      '����ʧ��---------------
                                    Call ServeLandingPointFault.push_back(ba)
                                    If tempSucceed > maxSucceed Then
                                          maxSucceed = tempSucceed
                                    End If
                                    tempSucceed = 0
'''����ʧ������������ж�
                                    If (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiA1)) _
                                    Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiA2)) _
                                    Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiB2)) _
                                    Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiB1)) Then
                                          Call ServeLandingMediumFault.push_back(ba)
                                          commit = "MediumFault"
                                          Call description(.Cells(i, aimColumn + 1), commit)
                                    ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiA1)) _
                                        Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiA2)) _
                                        Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiB2)) _
                                        Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiB1)) Then
                                          Call ServeLandingOuterFault.push_back(ba)
                                          commit = "OuterFault"
                                          Call description(.Cells(i, aimColumn + 1), commit)
                                    ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiA1)) _
                                        Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiA2)) _
                                        Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiB2)) _
                                        Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiB1)) Then
                                          Call ServeLandingInnerFault.push_back(ba)
                                          .Cells(j, 8) = "InnerFault"
                                    Else
                                          Call ServeLandingOtherFault.push_back(ba)
                                          commit = "OtherFault"
                                          Call description(.Cells(i, aimColumn + 1), commit)
                                    End If
'''
                              ElseIf .Cells(j, aimColumn) = "Let" Then        '��������-------------------
                                    Call ServeLetPoint.push_back(latestHit)
                              End If
                        ElseIf trainningMode = 2 Then     '�ӷ���ѵ��---------------------------------------------------------------------------
                              If .Cells(j, aimColumn) = "ServeIn" Then        '�ӷ���ι������----------------
                                    Call ReturnTotalPoint.push_back(ba)
                              ElseIf .Cells(j, aimColumn) = "Return" Then     '�ӷ������----------------
'''�ӷ��������ֻ�����ж�
                                    If ba.x > 0 Then
                                          If ba.y > -FOREBACK Then
                                                Call ReturnForehandTotal.push_back(ba)
                                                islatestHitForehand = True
                                          Else
                                                Call ReturnBackhandTotal.push_back(ba)
                                                islatestHitForehand = False
                                          End If
                                    Else
                                          If ba.y < FOREBACK Then
                                                Call ReturnForehandTotal.push_back(ba)
                                                islatestHitForehand = True
                                          Else
                                                Call ReturnBackhandTotal.push_back(ba)
                                                islatestHitForehand = False
                                          End If
                                    End If
'''
                                    Call latestHit.clone(ba)
                                    Call ReturnPoint.push_back(ba)
                              ElseIf .Cells(j, aimColumn) = "In" Then         '�ӷ����������----------------
'''�ӷ�����������Ŀ��������ж�
                                    Call ReturnTargetEasy.push_back(ba)
                                    Call ReturnInEasy.push_back(ba)
                                    commit = "Easy"
                                    If trainningLevel = 1 Then
                                          tempSucceed = tempSucceed + 1
      '''�ӷ��������������ж�
                                          If islatestHitForehand Then
                                                Call ReturnForehandIn.push_back(ba)
                                          Else
                                                Call ReturnBackhandIn.push_back(ba)
                                          End If
      '''
                                    End If
                                    If (latestHit.x > 0 And latestHit.y >= 0 And ba.isInside(ReturnTargetAreaNormalAntiA1)) Or _
                                       (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(ReturnTargetAreaNormalAntiA2)) Or _
                                       (latestHit.x < 0 And latestHit.y >= 0 And ba.isInside(ReturnTargetAreaNormalAntiB2)) Or _
                                       (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(ReturnTargetAreaNormalAntiB1)) Then
                                          Call ReturnTargetNormal.push_back(ba)
                                          commit = "Normal"
                                          If trainningLevel = 2 Then
                                                tempSucceed = tempSucceed + 1
      '''
                                                If islatestHitForehand Then
                                                      Call ReturnForehandIn.push_back(ba)
                                                Else
                                                      Call ReturnBackhandIn.push_back(ba)
                                                End If
      '''
                                          End If
                                    Else
            '''��������ɹ��ж�
                                          If trainningLevel = 2 Then
                                                If tempSucceed > maxSucceed Then
                                                      maxSucceed = tempSucceed
                                                End If
                                                tempSucceed = 0
            '''
                                          End If
                                          Call ReturnInNormal.push_back(ba)
                                    End If
                                    If (latestHit.x > 0 And latestHit.y >= 0 And ba.isInside(ReturnTargetAreaHardAntiA1)) Or _
                                       (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(ReturnTargetAreaHardAntiA2)) Or _
                                       (latestHit.x < 0 And latestHit.y >= 0 And ba.isInside(ReturnTargetAreaHardAntiB2)) Or _
                                       (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(ReturnTargetAreaHardAntiB1)) Then
                                          Call ReturnTargetHard.push_back(ba)
                                          commit = "Hard"
                                          If trainningLevel = 3 Then
                                                tempSucceed = tempSucceed + 1
      '''
                                                If islatestHitForehand Then
                                                      Call ReturnForehandIn.push_back(ba)
                                                Else
                                                      Call ReturnBackhandIn.push_back(ba)
                                                End If
      '''
                                          End If
                                    Else
                                          If trainningLevel = 3 Then
            '''
                                                If tempSucceed > maxSucceed Then
                                                      maxSucceed = tempSucceed
                                                End If
                                                tempSucceed = 0
            '''
                                          End If
                                          Call ReturnInHard.push_back(ba)
                                    End If
'''
                                    Call description(.Cells(i, aimColumn + 1), commit)
                              ElseIf .Cells(j, aimColumn) = "Out" Then        '�ӷ������ʧ��---------------
                                    Call ReturnOutPoint.push_back(ba)
            '''
                                    If tempSucceed > maxSucceed Then
                                          maxSucceed = tempSucceed
                                    End If
                                    tempSucceed = 0
            '''
                              End If
                        ElseIf trainningMode = 3 Then     '����ѵ��--------------------------------------------------------------------------------
                              If .Cells(j, aimColumn) = "Hit" Then            '����------------------
'''���������ֻ�����ж�
                                    If ba.x > 0 Then
                                          If ba.y > -FOREBACK Then
                                                Call HitForehandTotal.push_back(ba)
                                                islatestHitForehand = True
                                          Else
                                                Call HitBackhandTotal.push_back(ba)
                                                islatestHitForehand = False
                                          End If
                                    Else
                                          If ba.y < FOREBACK Then
                                                Call HitForehandTotal.push_back(ba)
                                                islatestHitForehand = True
                                          Else
                                                Call HitBackhandTotal.push_back(ba)
                                                islatestHitForehand = False
                                          End If
                                    End If
                                    Call latestHit.clone(ba)
                                    islatestHitTouchDown = False
'''
                              If .Cells(j, aimColumn) = "feedIn" Then            'ι������--------------
                                    Call HitPoint.push_back(ba)
                              ElseIf .Cells(j, aimColumn) = "feed" Then       'ι��------------------
                                    If Not islatestHitTouchDown Then
                                          Call HitBeingVolleyPoint.push_back(ba)
                                    End If
                              ElseIf .Cells(j, aimColumn) = "In" Then         '����ɹ�-----------------
                                    islatestHitTouchDown = True
                                    roundCount = roundCount + 1
'''������������Ŀ�������ж�
                                    Call HitTargetEasy.push_back(ba)
                                    Call HitInEasy.push_back(ba)
                                    commit = "Easy"
                                    If trainningLevel = 1 Then
                                          tempSucceed = tempSucceed + 1
      '''���������������ж�
                                          If islatestHitForehand Then
                                                Call HitForehandIn.push_back(ba)
                                          Else
                                                Call HitBackhandIn.push_back(ba)
                                          End If
      '''
                                    End If
                                    If ba.isInside(HitTargetAreaNormal) Then
                                          Call HitTargetNormal.push_back(ba)
                                          commit = "Normal"
                                          If trainningLevel = 2 Then
                                                tempSucceed = tempSucceed + 1
      '''���������������ж�
                                                If islatestHitForehand Then
                                                      Call HitForehandIn.push_back(ba)
                                                Else
                                                      Call HitBackhandIn.push_back(ba)
                                                End If
      '''
                                          End If
                                    Else
                                          If trainningLevel = 2 Then
                                                If tempSucceed > maxSucceed Then
                                                      maxSucceed = tempSucceed
                                                End If
                                                tempSucceed = 0
                                          End If
                                          Call HitInNormal.push_back(ba)
                                    End If
                                    If ba.isInside(HitTargetAreaHard) Then
                                          Call HitTargetHard.push_back(ba)
                                          commit = "Hard"
                                          If trainningLevel = 3 Then
                                                tempSucceed = tempSucceed + 1
      '''���������������ж�
                                                If islatestHitForehand Then
                                                      Call HitForehandIn.push_back(ba)
                                                Else
                                                      Call HitBackhandIn.push_back(ba)
                                                End If
      '''
                                          End If
                                    Else
                                          If trainningLevel = 3 Then
                                                If tempSucceed > maxSucceed Then
                                                      maxSucceed = tempSucceed
                                                End If
                                                tempSucceed = 0
                                          End If
                                          Call HitInHard.push_back(ba)
                                    End If
'''
                                    Call description(.Cells(i, aimColumn + 1), commit)
                              ElseIf .Cells(j, aimColumn) = "Out" Then        '����ʧ��------------------
'''������г����ж�
                                    If roundCount >= 10 Then
                                          longRound = longRound + 1
                                    ElseIf roundCount >= 6 Then
                                          middleRound = middleRound + 1
                                    ElseIf roundCount >= 2 Then
                                          shortRound = shortRound + 1
                                    End If
'''
                                    roundCount = 0
                                    islatestHitTouchDown = True
                                    Call HitOutPoint.push_back(ba)
                                    If tempSucceed > maxSucceed Then
                                          maxSucceed = tempSucceed
                                    End If
                                    tempSucceed = 0
                              End If
                        Else
                              commit = "?????????????????????????????????????????????"
                              Call description(.Cells(i, aimColumn + 1), commit)
                        End If
                  End If
            Next
            If roundCount >= 10 Then
                  longRound = longRound + 1
            ElseIf roundCount >= 6 Then
                  middleRound = middleRound + 1
            ElseIf roundCount >= 2 Then
                  shortRound = shortRound + 1
            End If
            If tempSucceed > maxSucceed Then
                  maxSucceed = tempSucceed
            End If
      End With
      
      '''''''''''''''''''���ݳ���'''''''''''''''''''
      '''''''''''''''''''���ݳ���'''''''''''''''''''
      '''''''''''''''''''���ݳ���'''''''''''''''''''
      '''''''''''''''''''���ݳ���'''''''''''''''''''
      '''''''''''''''''''���ݳ���'''''''''''''''''''
      
      ''''''''''trainningCharts'''''''''''''''''''''''''''''''
      Dim rowNum%, colNum%
      rowNum = 1
      commit = IIf(trainningSide > 0, "A��", "B��") & "_" & _
               IIf(trainningMode = 1, "����", IIf(trainningMode = 2, "�ӷ���", "����")) & "_" & _
               IIf(trainningLevel = 1, "��", IIf(trainningLevel = 2, "��", "��")) & "_"
      With Sheets("trainningCoordinates")
            Dim tempVecball As New vecBall
            colNum = IIf(.Cells(1, 100).End(1) = "", 1, .Cells(1, 100).End(1).Column + 2)
            If trainningMode = 1 Then                 '����ѵ��-------------------------------------------------------------------------------
                  .Cells(rowNum, colNum) = commit & "�ܴ���"
                  rowNum = rowNum + 1
                  .Cells(rowNum, colNum) = ServePoint.count
                  If trainningSide = 1 Then
                        Sheets("trainningOtherDetails").Range("E8").Value = ServePoint.count
                  Else
                        Sheets("trainningOtherDetails").Range("F8").Value = ServePoint.count
                  End If
                  rowNum = rowNum + 1
                  If ServePoint.count > 0 Then
                        Call tempVecball.combine(ServePoint)
                        For i = 0 To tempVecball.count - 1
                              Call ba.clone(tempVecball.pop_back())
                              .Cells(rowNum, colNum) = ba.x
                              .Cells(rowNum, colNum + 1) = ba.y
                              rowNum = rowNum + 1
                        Next
                  End If
                  colNum = colNum + 2
                  rowNum = 1
                  With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlMedium
                  End With
                  If trainningLevel = 1 Then          '��----------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ServeTargetEasy.count
                        If trainningSide = 1 Then
                              Sheets("trainningOtherDetails").Range("E5").Value = ServeTargetEasy.count
                              Sheets("trainningOtherDetails").Range("E6").Value = ServeTargetEasy.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = ServeTargetEasy.count
                              Sheets("trainningOtherDetails").Range("F6").Value = ServeTargetEasy.count
                        End If
                        rowNum = rowNum + 1
                        If ServeTargetEasy.count > 0 Then
                              For i = 0 To ServeTargetEasy.count - 1
                                    Call ba.clone(ServeTargetEasy.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  ElseIf trainningLevel = 2 Then      '��-------------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ServeTargetNormal.count
                        If trainningSide = 1 Then
                              Sheets("trainningOtherDetails").Range("E5").Value = ServeTargetNormal.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = ServeTargetNormal.count
                        End If
                        rowNum = rowNum + 1
                        If ServeTargetNormal.count > 0 Then
                              For i = 0 To ServeTargetNormal.count - 1
                                    Call ba.clone(ServeTargetNormal.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                        .Cells(rowNum, colNum) = commit & "��������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ServeInNormal.count
                        If trainningSide = 1 Then
                              Sheets("trainningOtherDetails").Range("E6").Value = ServeInNormal.count
                        Else
                              Sheets("trainningOtherDetails").Range("F6").Value = ServeInNormal.count
                        End If
                        rowNum = rowNum + 1
                        If ServeInNormal.count > 0 Then
                              For i = 0 To ServeInNormal.count - 1
                                    Call ba.clone(ServeInNormal.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  ElseIf trainningLevel = 3 Then      '��-----------------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ServeTargetHard.count
                        If trainningSide = 1 Then
                              Sheets("trainningOtherDetails").Range("E5").Value = ServeTargetHard.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = ServeTargetHard.count
                        End If
                        rowNum = rowNum + 1
                        If ServeTargetHard.count > 0 Then
                              For i = 0 To ServeTargetHard.count - 1
                                    Call ba.clone(ServeTargetHard.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                        .Cells(rowNum, colNum) = commit & "��������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ServeInHard.count
                        If trainningSide = 1 Then
                              Sheets("trainningOtherDetails").Range("E6").Value = ServeInHard.count
                        Else
                              Sheets("trainningOtherDetails").Range("F6").Value = ServeInHard.count
                        End If
                        rowNum = rowNum + 1
                        If ServeInHard.count > 0 Then
                              For i = 0 To ServeInHard.count - 1
                                    Call ba.clone(ServeInHard.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  End If
                  .Cells(rowNum, colNum) = commit & "�������"
                  rowNum = rowNum + 1
                  .Cells(rowNum, colNum) = ServeLandingPointFault.count
                  If trainningSide = 1 Then
                        Sheets("trainningOtherDetails").Range("E7").Value = ServeLandingPointFault.count
                  Else
                        Sheets("trainningOtherDetails").Range("F7").Value = ServeLandingPointFault.count
                  End If
                  rowNum = rowNum + 1
                  If ServeLandingPointFault.count > 0 Then
                        Call tempVecball.combine(ServeLandingPointFault)
                        For i = 0 To tempVecball.count - 1
                              Call ba.clone(tempVecball.pop_back())
                              .Cells(rowNum, colNum) = ba.x
                              .Cells(rowNum, colNum + 1) = ba.y
                              rowNum = rowNum + 1
                        Next
                  End If
                  colNum = colNum + 2
                  rowNum = 1
                  With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlThick
                  End With
            ElseIf trainningMode = 2 Then                   '�ӷ���ѵ��-----------------------------------------------------------------
                  .Cells(rowNum, colNum) = commit & "�ܴ���"
                  rowNum = rowNum + 1
                  .Cells(rowNum, colNum) = ReturnTotalPoint.count
                  If trainningSide = 1 Then
                        Sheets("trainningOtherDetails").Range("E8").Value = ReturnTotalPoint.count
                  Else
                        Sheets("trainningOtherDetails").Range("F8").Value = ReturnTotalPoint.count
                  End If
                  rowNum = rowNum + 1
                  If ReturnTotalPoint.count > 0 Then
                        Call tempVecball.combine(ReturnTotalPoint)
                        For i = 0 To tempVecball.count - 1
                              Call ba.clone(tempVecball.pop_back())
                              .Cells(rowNum, colNum) = ba.x
                              .Cells(rowNum, colNum + 1) = ba.y
                              rowNum = rowNum + 1
                        Next
                  End If
                  colNum = colNum + 2
                  rowNum = 1
                  With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlMedium
                  End With
                  If trainningLevel = 1 Then          '��----------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ReturnTargetEasy.count
                        If trainningSide = 1 Then
                              Sheets("trainningOtherDetails").Range("E5").Value = ReturnTargetEasy.count
                              Sheets("trainningOtherDetails").Range("E6").Value = ReturnTargetEasy.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = ReturnTargetEasy.count
                              Sheets("trainningOtherDetails").Range("F6").Value = ReturnTargetEasy.count
                        End If
                        rowNum = rowNum + 1
                        If ReturnTargetEasy.count > 0 Then
                              For i = 0 To ReturnTargetEasy.count - 1
                                    Call ba.clone(ReturnTargetEasy.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  ElseIf trainningLevel = 2 Then      '��-------------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ReturnTargetNormal.count
                        If trainningSide = 1 Then           'Ŀ��
                              Sheets("trainningOtherDetails").Range("E5").Value = ReturnTargetNormal.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = ReturnTargetNormal.count
                        End If
                        rowNum = rowNum + 1
                        If ReturnTargetNormal.count > 0 Then
                              For i = 0 To ReturnTargetNormal.count - 1
                                    Call ba.clone(ReturnTargetNormal.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                        .Cells(rowNum, colNum) = commit & "��������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ReturnInNormal.count
                        If trainningSide = 1 Then           '����
                              Sheets("trainningOtherDetails").Range("E6").Value = ReturnInNormal.count
                        Else
                              Sheets("trainningOtherDetails").Range("F6").Value = ReturnInNormal.count
                        End If
                        rowNum = rowNum + 1
                        If ReturnInNormal.count > 0 Then
                              For i = 0 To ReturnInNormal.count - 1
                                    Call ba.clone(ReturnInNormal.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  ElseIf trainningLevel = 3 Then      '��----------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ReturnTargetHard.count
                        If trainningSide = 1 Then           'Ŀ��
                              Sheets("trainningOtherDetails").Range("E5").Value = ReturnTargetHard.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = ReturnTargetHard.count
                        End If
                        rowNum = rowNum + 1
                        If ReturnTargetHard.count > 0 Then
                              For i = 0 To ReturnTargetHard.count - 1
                                    Call ba.clone(ReturnTargetHard.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                        .Cells(rowNum, colNum) = commit & "��������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = ReturnInHard.count
                        If trainningSide = 1 Then           '����
                              Sheets("trainningOtherDetails").Range("E6").Value = ReturnInHard.count
                        Else
                              Sheets("trainningOtherDetails").Range("F6").Value = ReturnInHard.count
                        End If
                        rowNum = rowNum + 1
                        If ReturnInHard.count > 0 Then
                              For i = 0 To ReturnInHard.count - 1
                                    Call ba.clone(ReturnInHard.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  End If
                  .Cells(rowNum, colNum) = commit & "�������"
                  rowNum = rowNum + 1
                  .Cells(rowNum, colNum) = ReturnOutPoint.count
                  If trainningSide = 1 Then           '����
                        Sheets("trainningOtherDetails").Range("E7").Value = ReturnOutPoint.count
                  Else
                        Sheets("trainningOtherDetails").Range("F7").Value = ReturnOutPoint.count
                  End If
                  rowNum = rowNum + 1
                  If ReturnOutPoint.count > 0 Then
                        Call tempVecball.combine(ReturnOutPoint)
                        For i = 0 To tempVecball.count - 1
                              Call ba.clone(tempVecball.pop_back())
                              .Cells(rowNum, colNum) = ba.x
                              .Cells(rowNum, colNum + 1) = ba.y
                              rowNum = rowNum + 1
                        Next
                  End If
                  colNum = colNum + 2
                  rowNum = 1
                  With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlMedium
                  End With
                  .Cells(rowNum, colNum) = commit & "�ܻ�����"
                  rowNum = rowNum + 1
                  .Cells(rowNum, colNum) = ReturnPoint.count
                  rowNum = rowNum + 1
                  If ReturnPoint.count > 0 Then
                        Call tempVecball.combine(ReturnPoint)
                        For i = 0 To tempVecball.count - 1
                              Call ba.clone(tempVecball.pop_back())
                              .Cells(rowNum, colNum) = ba.x
                              .Cells(rowNum, colNum + 1) = ba.y
                              rowNum = rowNum + 1
                        Next
                  End If
                  colNum = colNum + 2
                  rowNum = 1
                  With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlThick
                  End With
            ElseIf trainningMode = 3 Then                   '����ѵ��-----------------------------------------------------------------------------
                  .Cells(rowNum, colNum) = commit & "�ܴ���"
                  rowNum = rowNum + 1
                  .Cells(rowNum, colNum) = HitPoint.count
                  If trainningSide = 1 Then           '�ܴ���
                        Sheets("trainningOtherDetails").Range("E8").Value = HitPoint.count
                  Else
                        Sheets("trainningOtherDetails").Range("F8").Value = HitPoint.count
                  End If
                  rowNum = rowNum + 1
                  If HitPoint.count > 0 Then
                        Call tempVecball.combine(HitPoint)
                        For i = 0 To tempVecball.count - 1
                              Call ba.clone(tempVecball.pop_back())
                              .Cells(rowNum, colNum) = ba.x
                              .Cells(rowNum, colNum + 1) = ba.y
                              rowNum = rowNum + 1
                        Next
                  End If
                  colNum = colNum + 2
                  rowNum = 1
                  With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlMedium
                  End With
                  If trainningLevel = 1 Then          '��--------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = HitTargetEasy.count
                        If trainningSide = 1 Then           '��
                              Sheets("trainningOtherDetails").Range("E5").Value = HitTargetEasy.count
                              Sheets("trainningOtherDetails").Range("E6").Value = HitTargetEasy.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = HitTargetEasy.count
                              Sheets("trainningOtherDetails").Range("F6").Value = HitTargetEasy.count
                        End If
                        rowNum = rowNum + 1
                        If HitTargetEasy.count > 0 Then
                              For i = 0 To HitTargetEasy.count - 1
                                    Call ba.clone(HitTargetEasy.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  ElseIf trainningLevel = 2 Then      '��----------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = HitTargetNormal.count
                        If trainningSide = 1 Then           'Ŀ��
                              Sheets("trainningOtherDetails").Range("E5").Value = HitTargetNormal.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = HitTargetNormal.count
                        End If
                        rowNum = rowNum + 1
                        If HitTargetNormal.count > 0 Then
                              For i = 0 To HitTargetNormal.count - 1
                                    Call ba.clone(HitTargetNormal.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                        .Cells(rowNum, colNum) = commit & "��������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = HitInNormal.count
                        If trainningSide = 1 Then           '����
                              Sheets("trainningOtherDetails").Range("E6").Value = HitInNormal.count
                        Else
                              Sheets("trainningOtherDetails").Range("F6").Value = HitInNormal.count
                        End If
                        rowNum = rowNum + 1
                        If HitInNormal.count > 0 Then
                              For i = 0 To HitInNormal.count - 1
                                    Call ba.clone(HitInNormal.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  ElseIf trainningLevel = 3 Then      '��---------------------------------------------
                        .Cells(rowNum, colNum) = commit & "Ŀ������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = HitTargetHard.count
                        If trainningSide = 1 Then           'Ŀ��
                              Sheets("trainningOtherDetails").Range("E5").Value = HitTargetHard.count
                        Else
                              Sheets("trainningOtherDetails").Range("F5").Value = HitTargetHard.count
                        End If
                        rowNum = rowNum + 1
                        If HitTargetHard.count > 0 Then
                              For i = 0 To HitTargetHard.count - 1
                                    Call ba.clone(HitTargetHard.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                        .Cells(rowNum, colNum) = commit & "��������"
                        rowNum = rowNum + 1
                        .Cells(rowNum, colNum) = HitInHard.count
                        If trainningSide = 1 Then           '����
                              Sheets("trainningOtherDetails").Range("E6").Value = HitInHard.count
                        Else
                              Sheets("trainningOtherDetails").Range("F6").Value = HitInHard.count
                        End If
                        rowNum = rowNum + 1
                        If HitInHard.count > 0 Then
                              For i = 0 To HitInHard.count - 1
                                    Call ba.clone(HitInHard.pop_back())
                                    .Cells(rowNum, colNum) = ba.x
                                    .Cells(rowNum, colNum + 1) = ba.y
                                    rowNum = rowNum + 1
                              Next
                        End If
                        colNum = colNum + 2
                        rowNum = 1
                        With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                              .Borders(xlEdgeLeft).LineStyle = xlContinuous
                              .Borders(xlEdgeLeft).Weight = xlMedium
                        End With
                  End If
                  .Cells(rowNum, colNum) = commit & "�������"
                  rowNum = rowNum + 1
                  .Cells(rowNum, colNum) = HitOutPoint.count
                  If trainningSide = 1 Then           '����
                        Sheets("trainningOtherDetails").Range("E7").Value = HitOutPoint.count
                  Else
                        Sheets("trainningOtherDetails").Range("F7").Value = HitOutPoint.count
                  End If
                  rowNum = rowNum + 1
                  If HitOutPoint.count > 0 Then
                        Call tempVecball.combine(HitOutPoint)
                        For i = 0 To tempVecball.count - 1
                              Call ba.clone(tempVecball.pop_back())
                              .Cells(rowNum, colNum) = ba.x
                              .Cells(rowNum, colNum + 1) = ba.y
                              rowNum = rowNum + 1
                        Next
                  End If
                  colNum = colNum + 2
                  rowNum = 1
                  With .Range(.Cells(rowNum, colNum), .Cells(rowNum + 1, colNum))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlThick
                  End With
            End If
      End With
                
      '''''''''''''''''''''''''trainningOtherDetails'''''''''''''''''''''''''''''
      
      With Sheets("trainningOtherDetails")
            .Columns("E:F").HorizontalAlignment = xlCenter
            .Columns("M:N").HorizontalAlignment = xlCenter
            .Columns("I:J").HorizontalAlignment = xlCenter
            .Columns("Q:R").HorizontalAlignment = xlCenter

                  
            .Range("E3").Value = "A"
            .Range("F3").Value = "B"
            .Range("D4").Value = "��    ��"
            .Range("D5").Value = "Ŀ������"
            .Range("D6").Value = "��������"
            .Range("D7").Value = "��    ��"
            .Range("D8").Value = "�ܴ���"
            .Range("D9").Value = "������ɹ�"
            .Range("D10").Value = "���ڳɹ���"
            .Range("D11").Value = "Ŀ��ɹ���"
            If trainningSide = 1 Then
                  .Range("E9").Value = maxSucceed                            'A������ɹ�
                  .Range("E10").FormulaR1C1 = "=IfError(INT(100*(R6C5+R5C5)/R8C5)/100, 0)"      'A���ڳɹ���
                  .Range("E10").NumberFormatLocal = "0%"
                  .Range("E11").FormulaR1C1 = "=IfError(INT(100*R5C5/R8C5)/100, 0)"             'AĿ��ɹ���
                  .Range("E11").NumberFormatLocal = "0%"
                  .Range("E4").Value = IIf(trainningLevel = 1, .Range("E5") * 20, .Range("E5") * 20 + .Range("E6") * 10)
            Else
                  .Range("F9").Value = maxSucceed                            'B������ɹ�
                  .Range("F10").FormulaR1C1 = "=IfError(INT(100*��R6C6+R5C6)/R8C6)/100, 0)"     'B���ڳɹ���
                  .Range("F10").NumberFormatLocal = "0%"
                  .Range("F11").FormulaR1C1 = "=IfError(INT(100*R5C6/R8C6)/100, 0)"             'BĿ��ɹ���
                  .Range("F11").NumberFormatLocal = "0%"
                  .Range("F4").Value = IIf(trainningLevel = 1, .Range("F5") * 20, .Range("F5") * 20 + .Range("F6") * 10)
            End If
            
            .Cells(3, 9).Value = "A"
            .Cells(3, 10).Value = "B"
            .Cells(4, 8).Value = "�����ڽ�in"
            .Cells(5, 8).Value = "������·in"
            .Cells(6, 8).Value = "�������in"
            .Cells(7, 8).Value = "����ʧ��"
            .Cells(8, 8).Value = "�ڽ�Fault"
            .Cells(9, 8).Value = "��·Fault"
            .Cells(10, 8).Value = "���Fault"
            .Cells(11, 8).Value = "����Fault"
            .Cells(12, 8).Value = "�ڽǳɹ���"
            .Cells(13, 8).Value = "��·�ɹ���"
            .Cells(14, 8).Value = "��ǳɹ���"
            .Cells(16, 8).Value = "��   ��"
            .Cells(17, 8).Value = "�����ܴ���"
            
            .Cells(3, 13).Value = "A"
            .Cells(3, 14).Value = "B"
            .Cells(4, 12).Value = "�ӷ�������in"
            .Cells(5, 12).Value = "�ӷ�����in"
            .Cells(6, 12).Value = "�ӷ���out"
            .Cells(7, 12).Value = "�ӷ�������"
            .Cells(8, 12).Value = "�ӷ�����"
            .Cells(9, 12).Value = "�ӷ����ܴ���"
            .Cells(10, 12).Value = "�ӷ���ɹ���"
            
            .Cells(3, 17).Value = "A"
            .Cells(3, 18).Value = "B"
            .Cells(4, 16).Value = "��������in"
            .Cells(5, 16).Value = "������in"
            .Cells(6, 16).Value = "����out"
            .Cells(7, 16).Value = "��������"
            .Cells(8, 16).Value = "������"
            .Cells(9, 16).Value = "�����ܴ���"
            .Cells(10, 16).Value = "�������ֳɹ���"
            .Cells(11, 16).Value = "�����ֳɹ���"
            .Cells(13, 16).Value = "��   ��"
            .Cells(14, 16).Value = "��   ��"
            .Cells(15, 16).Value = "��   ��"
                  
            If trainningMode = 1 Then
                  If trainningSide = 1 Then
                        .Cells(4, 9).Value = ServeLandingInner.count          '�����ڽ�in
                        .Cells(5, 9).Value = ServeLandingMedium.count         '������·in
                        .Cells(6, 9).Value = ServeLandingOuter.count          '�������in
                        .Cells(7, 9).Value = ServeLandingPointFault.count     '����ʧ��
                        .Cells(8, 9).Value = ServeLandingInnerFault.count     '�ڽ�Fault
                        .Cells(9, 9).Value = ServeLandingMediumFault.count    '��·Fault
                        .Cells(10, 9).Value = ServeLandingOuterFault.count    '���Fault
                        .Cells(11, 9).Value = ServeLandingOtherFault.count    '����Fault
                        .Cells(12, 9).FormulaR1C1 = "=IfError(INT(100*R4C9/(R4C9 +R8C9))/100, 0)"           '�ڽǳɹ���
                        .Cells(12, 9).NumberFormatLocal = "0%"
                        .Cells(13, 9).FormulaR1C1 = "=IfError(INT(100*R5C9/(R5C9 +R9C9))/100, 0)"           '��·�ɹ���
                        .Cells(13, 9).NumberFormatLocal = "0%"
                        .Cells(14, 9).FormulaR1C1 = "=IfError(INT(100*R6C9/(R6C9 +R10C9))/100, 0)"          '��ǳɹ���
                        .Cells(14, 9).NumberFormatLocal = "0%"
                        .Cells(16, 9).Value = ServeLetPoint.count             '�����������
                        .Cells(17, 9).Value = ServePoint.count                '�����ܴ���
                  Else
                        .Cells(4, 10).Value = ServeLandingInner.count          '�����ڽ�in
                        .Cells(5, 10).Value = ServeLandingMedium.count         '������·in
                        .Cells(6, 10).Value = ServeLandingOuter.count          '�������in
                        .Cells(7, 10).Value = ServeLandingPointFault.count     '����ʧ��
                        .Cells(8, 10).Value = ServeLandingInnerFault.count     '�ڽ�Fault
                        .Cells(9, 10).Value = ServeLandingMediumFault.count    '��·Fault
                        .Cells(10, 10).Value = ServeLandingOuterFault.count    '���Fault
                        .Cells(11, 10).Value = ServeLandingOtherFault.count    '����Fault
                        
                        .Cells(12, 10).FormulaR1C1 = "=IfError(INT(100*R4C10/(R4C10 +R8C10))/100, 0)"       '�ڽǳɹ���
                        .Cells(12, 10).NumberFormatLocal = "0%"
                        .Cells(13, 10).FormulaR1C1 = "=IfError(INT(100*R5C10/(R5C10 +R9C10))/100, 0)"       '��·�ɹ���
                        .Cells(13, 10).NumberFormatLocal = "0%"
                        .Cells(14, 10).FormulaR1C1 = "=IfError(INT(100*R6C10/(R6C10 +R10C10))/100, 0)"      '��ǳɹ���
                        .Cells(14, 10).NumberFormatLocal = "0%"
                        .Cells(16, 10).Value = ServeLetPoint.count             '�����������
                        .Cells(17, 10).Value = ServePoint.count                '�����ܴ���
                  End If
            ElseIf trainningMode = 2 Then
                  If trainningSide = 1 Then
                        .Cells(4, 13).Value = ReturnForehandIn.count          '�ӷ�������in
                        .Cells(5, 13).Value = ReturnBackhandIn.count          '�ӷ�����in
                        .Cells(6, 13).Value = ReturnOutPoint.count            '�ӷ���out
                        .Cells(7, 13).Value = ReturnForehandTotal.count       '�ӷ�������
                        .Cells(8, 13).Value = ReturnBackhandTotal.count       '�ӷ�����
                        .Cells(9, 13).Value = ReturnPoint.count               '�ӷ����ܴ���
                        .Cells(10, 13).FormulaR1C1 = "=IfError(INT(100*(R4C13 +R5C13)/R9C13)/100, 0)"       '�ӷ���ɹ���
                        .Cells(10, 13).NumberFormatLocal = "0%"
                  Else
                        .Cells(4, 14).Value = ReturnForehandIn.count          '�ӷ�������in
                        .Cells(5, 14).Value = ReturnBackhandIn.count          '�ӷ�����in
                        .Cells(6, 14).Value = ReturnOutPoint.count            '�ӷ���out
                        .Cells(7, 14).Value = ReturnForehandTotal.count       '�ӷ�������
                        .Cells(8, 14).Value = ReturnBackhandTotal.count       '�ӷ�����
                        .Cells(9, 14).Value = ReturnPoint.count               '�ӷ����ܴ���
                        .Cells(10, 14).FormulaR1C1 = "=IfError(INT(100*(R4C14 +R5C14)/R9C14)/100, 0)"       '�ӷ���ɹ���
                        .Cells(10, 14).NumberFormatLocal = "0%"
                  End If
            Else
                  If trainningSide = 1 Then
                        .Cells(4, 17).Value = HitForehandIn.count          '��������in
                        .Cells(5, 17).Value = HitBackhandIn.count          '������in
                        .Cells(6, 17).Value = HitOutPoint.count            '����out
                        .Cells(7, 17).Value = HitForehandTotal.count       '��������
                        .Cells(8, 17).Value = HitBackhandTotal.count       '������
                        .Cells(9, 17).Value = HitPoint.count               '�����ܴ���
                        .Cells(10, 17).FormulaR1C1 = "=IfError(INT(100*R4C17/R7C17)/100, 0)"       '�������ֳɹ���
                        .Cells(10, 17).NumberFormatLocal = "0%"
                        .Cells(11, 17).FormulaR1C1 = "=IfError(INT(100*R5C17/R8C17)/100, 0)"       '�������ֳɹ���
                        .Cells(11, 17).NumberFormatLocal = "0%"
                        .Cells(13, 17).Value = shortRound               '��   ��
                        .Cells(14, 17).Value = middleRound              '��   ��
                        .Cells(15, 17).Value = longRound                '��   ��
                  Else
                        .Cells(4, 18).Value = HitForehandIn.count          '��������in
                        .Cells(5, 18).Value = HitBackhandIn.count          '������in
                        .Cells(6, 18).Value = HitOutPoint.count            '����out
                        .Cells(7, 18).Value = HitForehandTotal.count       '��������
                        .Cells(8, 18).Value = HitBackhandTotal.count       '������
                        .Cells(9, 18).Value = HitPoint.count               '�����ܴ���
                        .Cells(10, 18).FormulaR1C1 = "=IfError(INT(100*R4C18/R7C18)/100, 0)"       '�������ֳɹ���
                        .Cells(10, 18).NumberFormatLocal = "0%"
                        .Cells(11, 18).FormulaR1C1 = "=IfError(INT(100*R5C18/R8C18)/100, 0)"       '�������ֳɹ���
                        .Cells(11, 18).NumberFormatLocal = "0%"
                        .Cells(13, 18).Value = shortRound               '��   ��
                        .Cells(14, 18).Value = middleRound              '��   ��
                        .Cells(15, 18).Value = longRound                '��   ��
                  End If
            End If
            
            
      End With
End Sub

Private Sub description(rng As Range, words As String)
      rng.Value = words
End Sub




