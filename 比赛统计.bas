Attribute VB_Name = "ģ��4"
Option Explicit

Const VOLLEY As Integer = 100
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''����ͳ�Ʋ���'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub raceAnalysis()
      Dim aim As Range
      Dim i%, j%, k%, m%, FOREBACK%
      
      Dim ba As New ball
      
      Dim latestLandingIn As New ball           '��һ������㣬���ڼ�¼��ǰ�򲢲�����㵫�ǻغϽ���(��ACE������ʤ��)ʱ��������Ϣ
      Dim latestHit As New ball                 '��һ���������������(������߻���)
      Dim intervalHit As New ball               '���ϸ��������������(������߻���)�����ڼ�¼��ͳ�ƽӷ�����߻����Ѷ�ʱ�Է�����λ�ò���
      
      Dim Flag%, latestHitStat%, serveFlag%, errorCount%, roundCount%, latestHitCommit$
      'flag���ڱ�ʶ����ACE������ʤ�ֵĿ��ܣ���ǰΪ���Inʱflag=1,Ϊ��������(���򣬻���)ʱflag=0
      'latestHitStat���ڴ��ݱ��λ���(+ or -)���ǻ����¼����Ƿ����¼�(1 or 2)��������߼�
      'serveFlag������ʶ���غϷ���(+ or -)�Լ�һ��(1)����(2)��������ʱ�ı�
      'errorCount������¼���˶��ٸ���
      'roundCount������¼���غ�һ�����˶��ٸ����أ��Խӷ���(A+ B-)�Ļ�����Ϊ׼
      'latestHitCommit��¼���һ�λ���������������������ͨ�ػ�("hitBack")���ǽӷ���("return")
      Dim A_bout%, B_bout%, A_game%, B_game%, gameMode%
      
      Dim A1stServePoint As New vecBall                           'A_һ�������
      Dim A1stServeLandingPointOtherFault As New vecBall          'A_һ��ʧ���������
      Dim A1stServeLandingPointInnerFault As New vecBall          'A_һ��_�ڽ�ʧ�����
      Dim A1stServeLandingPointMediumFault As New vecBall         'A_һ��_��·ʧ�����
      Dim A1stServeLandingPointOuterFault As New vecBall          'A_һ��_���ʧ�����
      Dim A1stServeLandingPointInner As New vecBall               'A_һ�����_��
      Dim A1stServeLandingPointMedium As New vecBall              'A_һ�����_��
      Dim A1stServeLandingPointOuter As New vecBall               'A_һ�����_��
      Dim A1stServeLet As New vecBall                             'A_һ������
      Dim A2ndServeLet As New vecBall                             'A_��������
      Dim A2ndServePoint As New vecBall                           'A_���������
      Dim A2ndServeLandingPointOtherFault As New vecBall          'A_˫���������
      Dim A2ndServeLandingPointInnerFault As New vecBall          'A_����_�ڽ�ʧ�����
      Dim A2ndServeLandingPointMediumFault As New vecBall         'A_����_��·ʧ�����
      Dim A2ndServeLandingPointOuterFault As New vecBall          'A_����_���ʧ�����
      Dim A2ndServeLandingPointInner As New vecBall               'A_�������_��
      Dim A2ndServeLandingPointMedium As New vecBall              'A_�������_��
      Dim A2ndServeLandingPointOuter As New vecBall               'A_�������_��
      Dim AReturnPoint As New vecBall                             'A_�ӷ�������
      Dim AReturnLandingPointEasy As New vecBall                  'A_�ӷ������_��
      Dim AReturnLandingPointNormal As New vecBall                'A_�ӷ������_��
      Dim AReturnLandingPointHard As New vecBall                  'A_�ӷ������_��
      Dim AReturnLandingPointFault As New vecBall                 'A_�ӷ���ʧ�����
      Dim AReturnBeingVolleyPoint As New vecBall                  'A_�ӷ��򱻶Է��ػ�
      Dim AHitBeingVolleyPoint As New vecBall                     'A_���򱻶Է��ػ�
      Dim AHitPoint As New vecBall                                'A_�����
      Dim AHitLandingPointEasy As New vecBall                     'A_�������_��
      Dim AHitLandingPointNormal As New vecBall                   'A_�������_��
      Dim AHitLandingPointHard As New vecBall                     'A_�������_��
      Dim AHitLandingPointFault As New vecBall                    'A_����ʧ�����
      Dim ANetNeerByPoint As New vecBall                             'A_���������
      Dim ANetNeerByWin As New vecBall                           'A_��ǰ�÷�
'      Dim AAce As New vecBall                                     'A_ace
      Dim A1stServeAce As New vecBall                             'A_һ��ace
      Dim A2ndServeAce As New vecBall                             'A_����ace
      Dim AWinner As New vecBall                                  'A_��ʤ��
      Dim ABreakPoint As New vecBall                              'A_�Ʒ���
      Dim ABreakSucceed As New vecBall                            'A_�Ʒ��÷�
      Dim AShortRoundWin As New vecBall                           'A_����ʤ���غ���
      Dim AMediumRoundWin As New vecBall                          'A_����ʤ���غ���
      Dim ALongRoundWin As New vecBall                            'A_����ʤ���غ���
      Dim A1stServeWin As New vecBall                             'A_һ���÷ֻغ���
      Dim A2ndServeWin As New vecBall                             'A_�����÷ֻغ���
      Dim ABoutWin As New vecBall                                 'A_�ܵ÷���
      
      Dim A1stServeInWithoutAce As New vecBall                    'Aһ����ACE�����       //Ϊ�����ͳ�Ƽ���ģ�����������
      Dim A2ndServeInWithoutAce As New vecBall                    'A������ACE�����
      Dim B1stServeInWithoutAce As New vecBall                    'Bһ����ACE�����
      Dim B2ndServeInWithoutAce As New vecBall                    'B������ACE�����
      
      Dim B1stServePoint As New vecBall                           'B_һ�������
      Dim B1stServeLandingPointOtherFault As New vecBall          'B_һ��ʧ���������
      Dim B1stServeLandingPointInnerFault As New vecBall          'B_һ��_�ڽ�ʧ�����
      Dim B1stServeLandingPointMediumFault As New vecBall         'B_һ��_��·ʧ�����
      Dim B1stServeLandingPointOuterFault As New vecBall          'B_һ��_���ʧ�����
      Dim B1stServeLandingPointInner As New vecBall               'B_һ�����_��
      Dim B1stServeLandingPointMedium As New vecBall              'B_һ�����_��
      Dim B1stServeLandingPointOuter As New vecBall               'B_һ�����_��
      Dim B1stServeLet As New vecBall                             'B_һ������
      Dim B2ndServeLet As New vecBall                             'B_��������
      Dim B2ndServePoint As New vecBall                           'B_���������
      Dim B2ndServeLandingPointOtherFault As New vecBall          'B_˫���������
      Dim B2ndServeLandingPointInnerFault As New vecBall          'B_����_�ڽ�ʧ�����
      Dim B2ndServeLandingPointMediumFault As New vecBall         'B_����_��·ʧ�����
      Dim B2ndServeLandingPointOuterFault As New vecBall          'B_����_���ʧ�����
      Dim B2ndServeLandingPointInner As New vecBall               'B_�������_��
      Dim B2ndServeLandingPointMedium As New vecBall              'B_�������_��
      Dim B2ndServeLandingPointOuter As New vecBall               'B_�������_��
      Dim BReturnPoint As New vecBall                             'B_�ӷ�������
      Dim BReturnLandingPointEasy As New vecBall                  'B_�ӷ������_��
      Dim BReturnLandingPointNormal As New vecBall                'B_�ӷ������_��
      Dim BReturnLandingPointHard As New vecBall                  'B_�ӷ������_��
      Dim BReturnLandingPointFault As New vecBall                 'B_�ӷ���ʧ�����
      Dim BReturnBeingVolleyPoint As New vecBall                  'B_�ӷ��򱻶Է��ػ�
      Dim BHitBeingVolleyPoint As New vecBall                     'B_���򱻶Է��ػ�
      Dim BHitPoint As New vecBall                                'B_�����
      Dim BHitLandingPointEasy As New vecBall                     'B_�������_��
      Dim BHitLandingPointNormal As New vecBall                   'B_�������_��
      Dim BHitLandingPointHard As New vecBall                     'B_�������_��
      Dim BHitLandingPointFault As New vecBall                    'B_����ʧ�����
      Dim BNetNeerByPoint As New vecBall                             'B_���������
      Dim BNetNeerByWin As New vecBall                           'B_��ǰ�÷�
'      Dim BAce As New vecBall                                     'B_ace
      Dim B1stServeAce As New vecBall                             'B_һ��ace
      Dim B2ndServeAce As New vecBall                             'B_����ace
      Dim BWinner As New vecBall                                  'B_��ʤ��
      Dim BBreakPoint As New vecBall                              'B_�Ʒ���
      Dim BBreakSucceed As New vecBall                            'B_�Ʒ��÷�
      Dim BShortRoundWin As New vecBall                           'B_����ʤ���غ���
      Dim BMediumRoundWin As New vecBall                          'B_����ʤ���غ���
      Dim BLongRoundWin As New vecBall                            'B_����ʤ���غ���
      Dim B1stServeWin As New vecBall                             'B_һ���÷ֻغ���
      Dim B2ndServeWin As New vecBall                             'B_�����÷ֻغ���
      Dim BBoutWin As New vecBall                                 'B_�ܵ÷���
      
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
      
      Dim normalAntiB1 As New multiZone                           '����λ��B1������ʱ�����ػ����Ŀ������ _��
      Dim hardAntiB1 As New multiZone                             '����λ��B1������ʱ�����ػ����Ŀ������ _��
      Dim normalAntiB2 As New multiZone                           '����λ��B2������ʱ�����ػ����Ŀ������ _��
      Dim hardAntiB2 As New multiZone                             '����λ��B2������ʱ�����ػ����Ŀ������ _��
      Dim normalAntiA1 As New multiZone                           '����λ��A1������ʱ�����ػ����Ŀ������ _��
      Dim hardAntiA1 As New multiZone                             '����λ��A1������ʱ�����ػ����Ŀ������ _��
      Dim normalAntiA2 As New multiZone                           '����λ��A2������ʱ�����ػ����Ŀ������ _��
      Dim hardAntiA2 As New multiZone                             '����λ��A2������ʱ�����ػ����Ŀ������ _��
      '��Ϊ�����Ѷ��������ĸ����ޱ˴˲���������ֻ��Ҫ���������������򣬵��ǻ����Ѷ��������ĸ����޽��棬����Ҫ�ֱ�����
      
      Dim tmpZone As New zone
      Dim tempMultiZone As New multiZone
      
      Range(Sheets("main").Columns(8), Sheets("main").Columns(11)).Clear
      If Sheets("main").Range("N3").Value = 1 Then                'cm����
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
            
            If Sheets("main").Range("M3").Value = 1 Then          '����
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-1188, 411, 0, 229)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, 411, -1006, -411)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, -361, 0, -411)
                  Call tempMultiZone.push_back(tmpZone)
                  Call hardAntiB1.clone(tempMultiZone)        '������B1������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call hardAntiB2.clone(tempMultiZone)        '������B2������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorY
                  Call hardAntiA1.clone(tempMultiZone)        '������A1������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call hardAntiA2.clone(tempMultiZone)        '������A2������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-1188, 411, 0, 229)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, 411, -789, -411)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, -361, 0, -411)
                  Call tempMultiZone.push_back(tmpZone)
                  Call normalAntiB1.clone(tempMultiZone)      '������B1������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call normalAntiB2.clone(tempMultiZone)      '������B2������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorY
                  Call normalAntiA1.clone(tempMultiZone)      '������A1������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call normalAntiA2.clone(tempMultiZone)      '������A2������ʱ���������Ѷ�Ϊ�е�Ŀ������
            Else      '˫��
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-1188, 548, 0, 307)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, 548, -1006, -483)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, -483, 0, -548)
                  Call tempMultiZone.push_back(tmpZone)
                  Call hardAntiB1.clone(tempMultiZone)        '������B1������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call hardAntiB2.clone(tempMultiZone)        '������B2������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorY
                  Call hardAntiA1.clone(tempMultiZone)        '������A1������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call hardAntiA2.clone(tempMultiZone)        '������A2������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-1188, 548, 0, 307)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, 548, -789, -483)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, -483, 0, -548)
                  Call tempMultiZone.push_back(tmpZone)
                  Call normalAntiB1.clone(tempMultiZone)      '������B1������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call normalAntiB2.clone(tempMultiZone)      '������B2������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorY
                  Call normalAntiA1.clone(tempMultiZone)      '������A1������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call normalAntiA2.clone(tempMultiZone)      '������A2������ʱ���������Ѷ�Ϊ�е�Ŀ������
            End If
            
            
      ElseIf Sheets("main").Range("N3").Value = 2 Then
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
            
            If Sheets("main").Range("M3").Value = 1 Then          '����
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-11887, 4115, 0, 2286)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, 4115, -10059, -4115)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, -3615, 0, -4115)
                  Call tempMultiZone.push_back(tmpZone)
                  Call hardAntiB1.clone(tempMultiZone)        '������B1������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call hardAntiB2.clone(tempMultiZone)        '������B2������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorY
                  Call hardAntiA1.clone(tempMultiZone)        '������A1������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call hardAntiA2.clone(tempMultiZone)        '������A2������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-11887, 4115, 0, 2286)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, 4115, -7890, -4115)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, -3615, 0, -4115)
                  Call tempMultiZone.push_back(tmpZone)
                  Call normalAntiB1.clone(tempMultiZone)      '������B1������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call normalAntiB2.clone(tempMultiZone)      '������B2������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorY
                  Call normalAntiA1.clone(tempMultiZone)      '������A1������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call normalAntiA2.clone(tempMultiZone)      '������A2������ʱ���������Ѷ�Ϊ�е�Ŀ������
            Else      '˫��
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-11887, 5487, 0, 3073)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, 5487, -10059, -4829)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, -4829, 0, -5487)
                  Call tempMultiZone.push_back(tmpZone)
                  Call hardAntiB1.clone(tempMultiZone)        '������B1������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call hardAntiB2.clone(tempMultiZone)        '������B2������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorY
                  Call hardAntiA1.clone(tempMultiZone)        '������A1������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call hardAntiA2.clone(tempMultiZone)        '������A2������ʱ���������Ѷ�Ϊ�ѵ�Ŀ������
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-11887, 5487, 0, 3073)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, 5487, -789, -4829)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, -4829, 0, -5487)
                  Call tempMultiZone.push_back(tmpZone)
                  Call normalAntiB1.clone(tempMultiZone)      '������B1������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call normalAntiB2.clone(tempMultiZone)      '������B2������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorY
                  Call normalAntiA1.clone(tempMultiZone)      '������A1������ʱ���������Ѷ�Ϊ�е�Ŀ������
                  Call tempMultiZone.mirrorX
                  Call normalAntiA2.clone(tempMultiZone)      '������A2������ʱ���������Ѷ�Ϊ�е�Ŀ������
            End If
      Else
            MsgBox "cm/mm��ֵ���ò���"
            Exit Sub
      End If
      
      latestHitStat = 0                   '��ʶ���������������Լ��ǻ����Ƿ���
      serveFlag = 0                       '��ʶһ������
      Flag = 0                            '��ʶ��û����ʤ�ֻ���ACE��Ŀ���
      A_bout = 0
      B_bout = 0
      A_game = 0
      B_game = 0
      roundCount = 0
      gameMode = IIf(Sheets("main").Range("Q15").Value = "", 6, Sheets("main").Range("Q15").Value)    '����ģʽĬ����������
      
      
      Dim AReturnForehandTotal As New vecBall
      Dim AReturnBackhandTotal As New vecBall
      Dim AHitForehandTotal As New vecBall
      Dim AHitBackhandTotal As New vecBall
      Dim BReturnForehandTotal As New vecBall
      Dim BReturnBackhandTotal As New vecBall
      Dim BHitForehandTotal As New vecBall
      Dim BHitBackhandTotal As New vecBall
      
      Dim AReturnForehandIn As New vecBall
      Dim AReturnBackhandIn As New vecBall
      Dim AHitForehandIn As New vecBall
      Dim AHitBackhandIn As New vecBall
      Dim BReturnForehandIn As New vecBall
      Dim BReturnBackhandIn As New vecBall
      Dim BHitForehandIn As New vecBall
      Dim BHitBackhandIn As New vecBall
      
      Dim isLastHitForehand As Boolean
      '��һ����������������ʣ�����true;���ֻ���false
      
      
      With Sheets("main")
            .Columns("h").Clear
            Set aim = .Range(.Cells(1, 7), .Cells(9999, 7).End(xlUp))
            For i = 1 To aim.Cells.count
                  If aim.Cells(i) <> "" Then
                        j = aim.Cells(i).Row
                        Call ba.init(.Cells(j, 1), .Cells(j, 2), .Cells(j, 3), .Cells(j, 4), .Cells(j, 5))
                        If .Cells(j, 6) Like "Error*" Then                          '����
                              errorCount = errorCount + 1
                        ElseIf .Cells(j, 6) = "firstServe" Then                     'һ��
                              If .Cells(j, 1) * .Cells(j, 7) > 0 Then   'A����
                                    serveFlag = 1
                                    latestHitStat = 2
                                    Call A1stServePoint.push_back(ba)
                                    .Cells(j, 8) = "A1stServePoint"
                              Else                                      'B����
                                    serveFlag = -1
                                    latestHitStat = -2
                                    Call B1stServePoint.push_back(ba)
                                    .Cells(j, 8) = "B1stServePoint"
                              End If
                              Flag = 0
                              Call latestHit.clone(ba)
                              latestHitCommit = .Cells(j, 6)
                        ElseIf .Cells(j, 6) = "secondServe" Then                    '����
                              If .Cells(j, 1) * .Cells(j, 7) > 0 Then   'A����
                                    serveFlag = 2
                                    latestHitStat = 2
                                    Call A2ndServePoint.push_back(ba)
                                    .Cells(j, 8) = "A2ndServePoint"
                              Else                                      'B����
                                    serveFlag = -2
                                    latestHitStat = -2
                                    Call B2ndServePoint.push_back(ba)
                                    .Cells(j, 8) = "B2ndServePoint"
                              End If
                              Flag = 0
                              Call latestHit.clone(ba)
                              latestHitCommit = .Cells(j, 6)
                        ElseIf .Cells(j, 6) Like "let*" Then                        '����
                              If serveFlag = 2 Then         'A����
                                    Call A2ndServeLet.push_back(ba)
                              ElseIf serveFlag = -2 Then    'B����
                                    Call B2ndServeLet.push_back(ba)
                              ElseIf serveFlag = 1 Then     'Aһ��
                                    Call A1stServeLet.push_back(ba)
                              ElseIf serveFlag = -1 Then     'Bһ��
                                    Call B1stServeLet.push_back(ba)
                              Else
                                    .Cells(j, 8) = "9999999999999999999999999999999999999999999"
                              End If
                        ElseIf .Cells(j, 6) = "firstServeIn" Then                   'һ���ɹ�
                              If .Cells(j, 1) * .Cells(j, 7) < 0 Then   '��ǰ���λ��Bѡ�����ڷ�λ
                                    If ba.isInside(inner) Then
                                          Call A1stServeLandingPointInner.push_back(ba)
                                          .Cells(j, 8) = "A1stServeLandingPointInner"
                                    ElseIf ba.isInside(medium) Then
                                          Call A1stServeLandingPointMedium.push_back(ba)
                                          .Cells(j, 8) = "A1stServeLandingPointMedium"
                                    ElseIf ba.isInside(outer) Then
                                          Call A1stServeLandingPointOuter.push_back(ba)
                                          .Cells(j, 8) = "A1stServeLandingPointOuter"
                                    Else
                                          .Cells(j, 8) = "1111111111111111111111111111111111111111111111"
                                    End If
                                    Call A1stServeInWithoutAce.push_back(ba)
                              Else                                      '��ǰ���λ��Aѡ�����ڷ�λ
                                    If ba.isInside(inner) Then
                                          Call B1stServeLandingPointInner.push_back(ba)
                                          .Cells(j, 8) = "B1stServeLandingPointInner"
                                    ElseIf ba.isInside(medium) Then
                                          Call B1stServeLandingPointMedium.push_back(ba)
                                          .Cells(j, 8) = "B1stServeLandingPointMedium"
                                    ElseIf ba.isInside(outer) Then
                                          Call B1stServeLandingPointOuter.push_back(ba)
                                          .Cells(j, 8) = "B1stServeLandingPointOuter"
                                    Else
                                          .Cells(j, 8) = "22222222222222222222222222222222222222222222222222"
                                    End If
                                    Call B1stServeInWithoutAce.push_back(ba)
                              End If
                              Call latestLandingIn.clone(ba)
                              Flag = 1
                        ElseIf .Cells(j, 6) = "secondServeIn" Then                  '�����ɹ�
                              If .Cells(j, 1) * .Cells(j, 7) < 0 Then
                                    If ba.isInside(inner) Then
                                          Call A2ndServeLandingPointInner.push_back(ba)
                                          .Cells(j, 8) = "A2ndServeLandingPointInner"
                                    ElseIf ba.isInside(medium) Then
                                          Call A2ndServeLandingPointMedium.push_back(ba)
                                          .Cells(j, 8) = "A2ndServeLandingPointMedium"
                                    ElseIf ba.isInside(outer) Then
                                          Call A2ndServeLandingPointOuter.push_back(ba)
                                          .Cells(j, 8) = "A2ndServeLandingPointOuter"
                                    Else
                                          .Cells(j, 8) = "3333333333333333333333333333333333333333333333333"
                                    End If
                                    Call A2ndServeInWithoutAce.push_back(ba)
                              Else
                                    If ba.isInside(inner) Then
                                          Call B2ndServeLandingPointInner.push_back(ba)
                                          .Cells(j, 8) = "B2ndServeLandingPointInner"
                                    ElseIf ba.isInside(medium) Then
                                          Call B2ndServeLandingPointMedium.push_back(ba)
                                          .Cells(j, 8) = "B2ndServeLandingPointMedium"
                                    ElseIf ba.isInside(outer) Then
                                          Call B2ndServeLandingPointOuter.push_back(ba)
                                          .Cells(j, 8) = "B2ndServeLandingPointOuter"
                                    Else
                                          .Cells(j, 8) = "4444444444444444444444444444444444444444444444444444"
                                    End If
                                    Call B2ndServeInWithoutAce.push_back(ba)
                              End If
                              Call latestLandingIn.clone(ba)
                              Flag = 1
                        ElseIf .Cells(j, 6) = "return" Then                         '�ӷ���
                              Flag = 0
                              Call intervalHit.clone(latestHit)
                              Call latestHit.clone(ba)
                              latestHitCommit = .Cells(j, 6)
                              If .Cells(j, 1) * .Cells(j, 7) > 0 Then   'A���ӷ���
                                    latestHitStat = 1
                                    Call AReturnPoint.push_back(ba)
                                    .Cells(j, 8) = "AReturnPoint"
                                    roundCount = 1
                                    If ba.x > 0 Then
                                          If ba.y > -FOREBACK Then
                                                Call AReturnForehandTotal.push_back(ba)
                                                isLastHitForehand = True
                                          Else
                                                Call AReturnBackhandTotal.push_back(ba)
                                                isLastHitForehand = False
                                          End If
                                    Else
                                          If ba.y < FOREBACK Then
                                                Call AReturnForehandTotal.push_back(ba)
                                                isLastHitForehand = True
                                          Else
                                                Call AReturnBackhandTotal.push_back(ba)
                                                isLastHitForehand = False
                                          End If
                                    End If
                              Else                                      'B���ӷ���
                                    latestHitStat = -1
                                    Call BReturnPoint.push_back(ba)
                                    .Cells(j, 8) = "BReturnPoint"
                                    roundCount = -1
                                    If ba.x > 0 Then
                                          If ba.y > -FOREBACK Then
                                                Call BReturnForehandTotal.push_back(ba)
                                                isLastHitForehand = True
                                          Else
                                                Call BReturnBackhandTotal.push_back(ba)
                                                isLastHitForehand = False
                                          End If
                                    Else
                                          If ba.y < FOREBACK Then
                                                Call BReturnForehandTotal.push_back(ba)
                                                isLastHitForehand = True
                                          Else
                                                Call BReturnBackhandTotal.push_back(ba)
                                                isLastHitForehand = False
                                          End If
                                    End If
                              End If
                        ElseIf .Cells(j, 6) = "hitBack" Then                        '����
                              '�����ǰ����Ϊ�ػ�������һ������ΪIN
                              If Flag = 0 Then
                                    If latestHitCommit = "return" Then   '�ӷ����ػ�
                                          If .Cells(j, 1) * .Cells(j, 7) > 0 Then   '��ǰΪA������
                                                Call BReturnBeingVolleyPoint.push_back(latestHit)
                                                If isLastHitForehand Then
                                                      Call BReturnForehandIn.push_back(latestHit)
                                                Else
                                                      Call BReturnBackhandIn.push_back(latestHit)
                                                End If
                                          Else                                      '��ǰΪB������
                                                Call AReturnBeingVolleyPoint.push_back(latestHit)
                                                If isLastHitForehand Then
                                                      Call AReturnForehandIn.push_back(latestHit)
                                                Else
                                                      Call AReturnBackhandIn.push_back(latestHit)
                                                End If
                                          End If
                                    ElseIf latestHitCommit = "hitBack" Then   '���򱻽ػ�
                                          If .Cells(j, 1) * .Cells(j, 7) > 0 Then   '��ǰΪA������
                                                Call BHitBeingVolleyPoint.push_back(latestHit)
                                                If isLastHitForehand Then
                                                      Call BHitForehandIn.push_back(latestHit)
                                                Else
                                                      Call BHitBackhandIn.push_back(latestHit)
                                                End If
                                          Else                                      '��ǰΪB������
                                                Call AHitBeingVolleyPoint.push_back(latestHit)
                                                If isLastHitForehand Then
                                                      Call AHitForehandIn.push_back(latestHit)
                                                Else
                                                      Call AHitBackhandIn.push_back(latestHit)
                                                End If
                                          End If
                                    End If
                              End If
                              '��ʼ����ǰ��������
                              If .Cells(j, 1) * .Cells(j, 7) > 0 Then   '��ǰΪA������
                                    latestHitStat = 1
                                    Call AHitPoint.push_back(ba)
                                    If Abs(.Cells(j, 1)) <= VOLLEY Then
                                          Call ANetNeerByPoint.push_back(ba)
                                          .Cells(j, 8) = "ANetNeerByPoint"
                                    Else
                                          .Cells(j, 8) = "AHitPoint"
                                    End If
                                    If roundCount > 0 Then
                                          roundCount = roundCount + 1
                                    End If
                                    If ba.x > 0 Then
                                          If ba.y > -FOREBACK Then
                                                Call AHitForehandTotal.push_back(ba)
                                                isLastHitForehand = True
                                          Else
                                                Call AHitBackhandTotal.push_back(ba)
                                                isLastHitForehand = False
                                          End If
                                    Else
                                          If ba.y < FOREBACK Then
                                                Call AHitForehandTotal.push_back(ba)
                                                isLastHitForehand = True
                                          Else
                                                Call AHitBackhandTotal.push_back(ba)
                                                isLastHitForehand = False
                                          End If
                                    End If
                              Else                                      '��ǰΪB������
                                    latestHitStat = -1
                                    Call BHitPoint.push_back(ba)
                                    If Abs(.Cells(j, 1)) <= VOLLEY Then
                                          Call BNetNeerByPoint.push_back(ba)
                                          .Cells(j, 8) = "BNetNeerByPoint"
                                    Else
                                          .Cells(j, 8) = "BHitPoint"
                                    End If
                                    If roundCount < 0 Then
                                          roundCount = roundCount - 1
                                    End If
                                    If ba.x > 0 Then
                                          If ba.y > -FOREBACK Then
                                                Call BHitForehandTotal.push_back(ba)
                                                isLastHitForehand = True
                                          Else
                                                Call BHitBackhandTotal.push_back(ba)
                                                isLastHitForehand = False
                                          End If
                                    Else
                                          If ba.y < FOREBACK Then
                                                Call BHitForehandTotal.push_back(ba)
                                                isLastHitForehand = True
                                          Else
                                                Call BHitBackhandTotal.push_back(ba)
                                                isLastHitForehand = False
                                          End If
                                    End If
                              End If
                              Call intervalHit.clone(latestHit)
                              Call latestHit.clone(ba)
                              latestHitCommit = .Cells(j, 6)
                              Flag = 0
                        ElseIf .Cells(j, 6) = "in" Then                             '����
                              Call latestLandingIn.clone(ba)
                              Flag = 1
                              If .Cells(j, 1) * .Cells(j, 7) < 0 Then     '��ǰ�����B������������
                                    If latestHitCommit = "return" Then
                                          Call AReturnLandingPointEasy.push_back(ba)
                                          .Cells(j, 8) = "AReturnLandingPointEasy"
                                          If isLastHitForehand Then
                                                Call AReturnForehandIn.push_back(latestHit)
                                          Else
                                                Call AReturnBackhandIn.push_back(latestHit)
                                          End If
                                    Else
                                          Call AHitLandingPointEasy.push_back(ba)
                                          .Cells(j, 8) = "AHitLandingPointEasy"
                                          If isLastHitForehand Then
                                                Call AHitForehandIn.push_back(latestHit)
                                          Else
                                                Call AHitBackhandIn.push_back(latestHit)
                                          End If
                                    End If
                                    If intervalHit.x > 0 Then
                                          If intervalHit.y > 0 Then
                                                If ba.isInside(normalAntiA1) Then
                                                      If latestHitCommit = "return" Then
                                                            Call AReturnLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "AReturnLandingPointNormal"
                                                      Else
                                                            Call AHitLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "AHitLandingPointNormal"
                                                      End If
                                                End If
                                                If ba.isInside(hardAntiA1) Then
                                                      If latestHitCommit = "return" Then
                                                            Call AReturnLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "AReturnLandingPointHard"
                                                      Else
                                                            Call AHitLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "AHitLandingPointHard"
                                                      End If
                                                End If
                                          Else
                                                If ba.isInside(normalAntiA2) Then
                                                      If latestHitCommit = "return" Then
                                                            Call AReturnLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "AReturnLandingPointNormal"
                                                      Else
                                                            Call AHitLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "AHitLandingPointNormal"
                                                      End If
                                                End If
                                                If ba.isInside(hardAntiA2) Then
                                                      If latestHitCommit = "return" Then
                                                            Call AReturnLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "AReturnLandingPointHard"
                                                      Else
                                                            Call AHitLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "AHitLandingPointHard"
                                                      End If
                                                End If
                                          End If
                                    Else
                                          If intervalHit.y > 0 Then
                                                If ba.isInside(normalAntiB2) Then
                                                      If latestHitCommit = "return" Then
                                                            Call AReturnLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "AReturnLandingPointNormal"
                                                      Else
                                                            Call AHitLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "AHitLandingPointNormal"
                                                      End If
                                                End If
                                                If ba.isInside(hardAntiB2) Then
                                                      If latestHitCommit = "return" Then
                                                            Call AReturnLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "AReturnLandingPointHard"
                                                      Else
                                                            Call AHitLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "AHitLandingPointHard"
                                                      End If
                                                End If
                                          Else
                                                If ba.isInside(normalAntiB1) Then
                                                      If latestHitCommit = "return" Then
                                                            Call AReturnLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "AReturnLandingPointNormal"
                                                      Else
                                                            Call AHitLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "AHitLandingPointNormal"
                                                      End If
                                                End If
                                                If ba.isInside(hardAntiB1) Then
                                                      If latestHitCommit = "return" Then
                                                            Call AReturnLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "AReturnLandingPointHard"
                                                      Else
                                                            Call AHitLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "AHitLandingPointHard"
                                                      End If
                                                End If
                                          End If
                                    End If
                              Else                    '��ǰ�����A����������
                                    If latestHitCommit = "return" Then
                                          Call BReturnLandingPointEasy.push_back(ba)
                                          .Cells(j, 8) = "BReturnLandingPointEasy"
                                          If isLastHitForehand Then
                                                Call BReturnForehandIn.push_back(latestHit)
                                          Else
                                                Call BReturnBackhandIn.push_back(latestHit)
                                          End If
                                    Else
                                          Call BHitLandingPointEasy.push_back(ba)
                                          .Cells(j, 8) = "BHitLandingPointEasy"
                                          If isLastHitForehand Then
                                                Call BHitForehandIn.push_back(latestHit)
                                          Else
                                                Call BHitBackhandIn.push_back(latestHit)
                                          End If
                                    End If
                                    If intervalHit.x > 0 Then
                                          If intervalHit.y > 0 Then
                                                If ba.isInside(normalAntiA1) Then
                                                      If latestHitCommit = "return" Then
                                                            Call BReturnLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "BReturnLandingPointNormal"
                                                      Else
                                                            Call BHitLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "BHitLandingPointNormal"
                                                      End If
                                                End If
                                                If ba.isInside(hardAntiA1) Then
                                                      If latestHitCommit = "return" Then
                                                            Call BReturnLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "BReturnLandingPointHard"
                                                      Else
                                                            Call BHitLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "BHitLandingPointHard"
                                                      End If
                                                End If
                                          Else
                                                If ba.isInside(normalAntiA2) Then
                                                      If latestHitCommit = "return" Then
                                                            Call BReturnLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "BReturnLandingPointNormal"
                                                      Else
                                                            Call BHitLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "BHitLandingPointNormal"
                                                      End If
                                                End If
                                                If ba.isInside(hardAntiA2) Then
                                                      If latestHitCommit = "return" Then
                                                            Call BReturnLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "BReturnLandingPointHard"
                                                      Else
                                                            Call BHitLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "BHitLandingPointHard"
                                                      End If
                                                End If
                                          End If
                                    Else
                                          If intervalHit.y > 0 Then
                                                If ba.isInside(normalAntiB2) Then
                                                      If latestHitCommit = "return" Then
                                                            Call BReturnLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "BReturnLandingPointNormal"
                                                      Else
                                                            Call BHitLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "BHitLandingPointNormal"
                                                      End If
                                                End If
                                                If ba.isInside(hardAntiB2) Then
                                                      If latestHitCommit = "return" Then
                                                            Call BReturnLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "BReturnLandingPointHard"
                                                      Else
                                                            Call BHitLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "BHitLandingPointHard"
                                                      End If
                                                End If
                                          Else
                                                If ba.isInside(normalAntiB1) Then
                                                      If latestHitCommit = "return" Then
                                                            Call BReturnLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "BReturnLandingPointNormal"
                                                      Else
                                                            Call BHitLandingPointNormal.push_back(ba)
                                                            .Cells(j, 8) = "BHitLandingPointNormal"
                                                      End If
                                                End If
                                                If ba.isInside(hardAntiB1) Then
                                                      If latestHitCommit = "return" Then
                                                            Call BReturnLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "BReturnLandingPointHard"
                                                      Else
                                                            Call BHitLandingPointHard.push_back(ba)
                                                            .Cells(j, 8) = "BHitLandingPointHard"
                                                      End If
                                                End If
                                          End If
                                    End If
                              End If
                        ElseIf .Cells(j, 6) = "fault,waitingForSecondServe" Or .Cells(j, 6) = "faultGuess,waitingForSecondServe" Then      'һ��ʧ��
                              If latestHitStat = 2 Then
                                    If (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiA1)) _
                                    Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiA2)) _
                                    Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiB2)) _
                                    Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiB1)) Then
                                          Call A1stServeLandingPointMediumFault.push_back(ba)
                                          .Cells(j, 8) = "A1stServeLandingPointMediumFault"
                                    ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiA1)) _
                                        Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiA2)) _
                                        Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiB2)) _
                                        Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiB1)) Then
                                          Call A1stServeLandingPointOuterFault.push_back(ba)
                                          .Cells(j, 8) = "A1stServeLandingPointOuterFault"
                                    ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiA1)) _
                                        Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiA2)) _
                                        Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiB2)) _
                                        Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiB1)) Then
                                          Call A1stServeLandingPointInnerFault.push_back(ba)
                                          .Cells(j, 8) = "A1stServeLandingPointInnerFault"
                                    Else
                                          Call A1stServeLandingPointOtherFault.push_back(ba)
                                          .Cells(j, 8) = "A1stServeLandingPointOtherFault"
                                    End If
                              ElseIf latestHitStat = -2 Then
                                    If (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiA1)) _
                                    Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiA2)) _
                                    Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiB2)) _
                                    Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiB1)) Then
                                          Call B1stServeLandingPointMediumFault.push_back(ba)
                                          .Cells(j, 8) = "B1stServeLandingPointMediumFault"
                                    ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiA1)) _
                                        Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiA2)) _
                                        Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiB2)) _
                                        Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiB1)) Then
                                          Call B1stServeLandingPointOuterFault.push_back(ba)
                                          .Cells(j, 8) = "B1stServeLandingPointOuterFault"
                                    ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiA1)) _
                                        Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiA2)) _
                                        Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiB2)) _
                                        Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiB1)) Then
                                          Call B1stServeLandingPointInnerFault.push_back(ba)
                                          .Cells(j, 8) = "B1stServeLandingPointInnerFault"
                                    Else
                                          Call B1stServeLandingPointOtherFault.push_back(ba)
                                          .Cells(j, 8) = "B1stServeLandingPointOtherFault"
                                    End If
                              Else
                                    .Cells(j, 8) = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB"
                              End If
                        ElseIf .Cells(j, 6) Like "*boutEnd" Then                                            '�غϽ���
                              If Flag = 1 Then                          '��һ�λ���ɹ����
                                    If latestHitStat = 2 Then           '�����һ�λ�����A������ ace
                                          A_bout = A_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) > 0, ba, A_bout, A_game, ABreakPoint, ABreakSucceed, _
                                          serveFlag, A1stServeWin, A2ndServeWin, ABoutWin, B_bout, B_game)
                                          If Abs(serveFlag) = 1 Then
                                                Call A1stServeAce.push_back(latestLandingIn)
                                                .Cells(j, 8) = "A1stServeAce"
                                                Call A1stServeInWithoutAce.pop_back
                                          ElseIf Abs(serveFlag) = 2 Then
                                                Call A2ndServeAce.push_back(latestLandingIn)
                                                .Cells(j, 8) = "A2ndServeAce"
                                                Call A2ndServeInWithoutAce.pop_back
                                          End If
                                    ElseIf latestHitStat = -2 Then      '�����һ�λ�����B������ ace
                                          B_bout = B_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) < 0, ba, B_bout, B_game, BBreakPoint, BBreakSucceed, _
                                          serveFlag, B1stServeWin, B2ndServeWin, BBoutWin, A_bout, A_game)
                                          If Abs(serveFlag) = 1 Then
                                                Call B1stServeAce.push_back(latestLandingIn)
                                                .Cells(j, 8) = "B1stServeAce"
                                                Call B1stServeInWithoutAce.pop_back
                                          ElseIf Abs(serveFlag) = 2 Then
                                                Call B2ndServeAce.push_back(latestLandingIn)
                                                .Cells(j, 8) = "B2ndServeAce"
                                                Call B2ndServeInWithoutAce.pop_back
                                          End If
                                    ElseIf latestHitStat = 1 Then       '�����һ�λ�����A������ winner
                                          A_bout = A_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) > 0, ba, A_bout, A_game, ABreakPoint, ABreakSucceed, _
                                          serveFlag, A1stServeWin, A2ndServeWin, ABoutWin, B_bout, B_game)
                                          Call AWinner.push_back(latestLandingIn)
                                          .Cells(j, 8) = "AWinner"
                                          If Abs(latestHit.x) <= VOLLEY Then
                                                Call ANetNeerByWin.push_back(ba)
                                          End If
                                          If Abs(roundCount) >= 10 Then
                                                Call ALongRoundWin.push_back(ba)
                                          ElseIf Abs(roundCount) >= 6 Then
                                                Call AMediumRoundWin.push_back(ba)
                                          ElseIf Abs(roundCount) >= 2 Then
                                                Call AShortRoundWin.push_back(ba)
                                          End If
                                    ElseIf latestHitStat = -1 Then      '�����һ�λ�����B������ winner
                                          B_bout = B_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) < 0, ba, B_bout, B_game, BBreakPoint, BBreakSucceed, _
                                          serveFlag, B1stServeWin, B2ndServeWin, BBoutWin, A_bout, A_game)
                                          Call BWinner.push_back(latestLandingIn)
                                          .Cells(j, 8) = "BWinner"
                                          If Abs(latestHit.x) <= VOLLEY Then
                                                Call BNetNeerByWin.push_back(ba)
                                          End If
                                          If Abs(roundCount) >= 10 Then
                                                Call BLongRoundWin.push_back(ba)
                                          ElseIf Abs(roundCount) >= 6 Then
                                                Call BMediumRoundWin.push_back(ba)
                                          ElseIf Abs(roundCount) >= 2 Then
                                                Call BShortRoundWin.push_back(ba)
                                          End If
                                    Else
                                          .Cells(j, 8) = "DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD"
                                    End If
                              Else                                      '��һ�λ���û�гɹ����
                                    If latestHitStat = -2 Then          '�����һ�λ�����B������,doubleFault
                                          If (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiA1)) _
                                          Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiA2)) _
                                          Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiB2)) _
                                          Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiB1)) Then
                                                Call B2ndServeLandingPointMediumFault.push_back(ba)
                                                .Cells(j, 8) = "B2ndServeLandingPointMediumFault"
                                          ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiA1)) _
                                              Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiA2)) _
                                              Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiB2)) _
                                              Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiB1)) Then
                                                Call B2ndServeLandingPointOuterFault.push_back(ba)
                                                .Cells(j, 8) = "B2ndServeLandingPointOuterFault"
                                          ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiA1)) _
                                              Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiA2)) _
                                              Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiB2)) _
                                              Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiB1)) Then
                                                Call B2ndServeLandingPointInnerFault.push_back(ba)
                                                .Cells(j, 8) = "B2ndServeLandingPointInnerFault"
                                          Else
                                                Call B2ndServeLandingPointOtherFault.push_back(ba)
                                                .Cells(j, 8) = "B2ndServeLandingPointOtherFault"
                                          End If
                                          A_bout = A_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) > 0, ba, A_bout, A_game, ABreakPoint, ABreakSucceed, _
                                          serveFlag, A1stServeWin, A2ndServeWin, ABoutWin, B_bout, B_game)
                                    ElseIf latestHitStat = 2 Then       '�����һ�λ�����A������,doubleFault
                                          If (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiA1)) _
                                          Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiA2)) _
                                          Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(mediumFaultAntiB2)) _
                                          Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(mediumFaultAntiB1)) Then
                                                Call A2ndServeLandingPointMediumFault.push_back(ba)
                                                .Cells(j, 8) = "A2ndServeLandingPointMediumFault"
                                          ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiA1)) _
                                              Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiA2)) _
                                              Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(outerFaultAntiB2)) _
                                              Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(outerFaultAntiB1)) Then
                                                Call A2ndServeLandingPointOuterFault.push_back(ba)
                                                .Cells(j, 8) = "A2ndServeLandingPointOuterFault"
                                          ElseIf (latestHit.x > 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiA1)) _
                                              Or (latestHit.x > 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiA2)) _
                                              Or (latestHit.x < 0 And latestHit.y > 0 And ba.isInside(innerFaultAntiB2)) _
                                              Or (latestHit.x < 0 And latestHit.y < 0 And ba.isInside(innerFaultAntiB1)) Then
                                                Call A2ndServeLandingPointInnerFault.push_back(ba)
                                                .Cells(j, 8) = "A2ndServeLandingPointInnerFault"
                                          Else
                                                Call A2ndServeLandingPointOtherFault.push_back(ba)
                                                .Cells(j, 8) = "A2ndServeLandingPointOtherFault"
                                          End If
                                          B_bout = B_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) < 0, ba, B_bout, B_game, BBreakPoint, BBreakSucceed, _
                                          serveFlag, B1stServeWin, B2ndServeWin, BBoutWin, A_bout, A_game)
                                    ElseIf latestHitStat = -1 Then      '�����һ�λ�����B������ fault
                                          A_bout = A_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) > 0, ba, A_bout, A_game, ABreakPoint, ABreakSucceed, _
                                          serveFlag, A1stServeWin, A2ndServeWin, ABoutWin, B_bout, B_game)
                                          If latestHitCommit = "return" Then        '��һ�λ���Ϊ�ӷ���
                                                Call BReturnLandingPointFault.push_back(latestHit)
                                                .Cells(j, 8) = "BReturnLandingPointFault"
                                          Else                                      '��һ�λ���Ϊ��ͨ����
                                                Call BHitLandingPointFault.push_back(latestHit)
                                                .Cells(j, 8) = "BHitLandingPointFault"
                                                If Abs(roundCount) >= 10 Then
                                                      Call ALongRoundWin.push_back(ba)
                                                ElseIf Abs(roundCount) >= 6 Then
                                                      Call AMediumRoundWin.push_back(ba)
                                                ElseIf Abs(roundCount) >= 2 Then
                                                      Call AShortRoundWin.push_back(ba)
                                                End If
                                          End If
                                          If Abs(intervalHit.x) <= VOLLEY Then
                                                Call ANetNeerByWin.push_back(ba)
                                          End If
                                    ElseIf latestHitStat = 1 Then       '�����һ�λ�����A������ fault
                                          B_bout = B_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) < 0, ba, B_bout, B_game, BBreakPoint, BBreakSucceed, _
                                          serveFlag, B1stServeWin, B2ndServeWin, BBoutWin, A_bout, A_game)
                                          If latestHitCommit = "return" Then        '��һ�λ���Ϊ�ӷ���
                                                Call AReturnLandingPointFault.push_back(latestHit)
                                                .Cells(j, 8) = "AReturnLandingPointFault"
                                          Else                                      '��һ�λ���Ϊ��ͨ����
                                                Call AHitLandingPointFault.push_back(latestHit)
                                                .Cells(j, 8) = "AHitLandingPointFault"
                                                If Abs(roundCount) >= 10 Then
                                                      Call BLongRoundWin.push_back(ba)
                                                ElseIf Abs(roundCount) >= 6 Then
                                                      Call BMediumRoundWin.push_back(ba)
                                                ElseIf Abs(roundCount) >= 2 Then
                                                      Call BShortRoundWin.push_back(ba)
                                                End If
                                          End If
                                          If Abs(intervalHit.x) <= VOLLEY Then
                                                Call BNetNeerByWin.push_back(ba)
                                          End If
                                    Else
                                          .Cells(j, 8) = "LLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLL"
                                    End If
                              End If
                              roundCount = 0
                        End If
                  End If
            Next i
      End With
      
      '''''''''''''''''''''''''''''''''''''���������������''''''''''''''''''''''''''''''''''''''''
      With Sheets("raceCoordinates")
            .Columns("A:T").Clear
            .Cells(1, 1) = "Aһ��ACE"
            .Cells(1, 3) = "Aһ����ACE"
            .Cells(1, 5) = "A����ACE"
            .Cells(1, 7) = "A������ACE"
            .Cells(1, 9) = "Bһ��ACE"
            .Cells(1, 11) = "Bһ����ACE"
            .Cells(1, 13) = "B����ACE"
            .Cells(1, 15) = "B������ACE"
            .Cells(1, 17) = "A�����"
            .Cells(1, 19) = "B�����"
            
            Dim bb As New vecBall
            Call bb.init
            Call bb.combine(A1stServeAce)                               'Aһ��ACE
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 1) = ba.x
                  .Cells(i + 2, 2) = ba.y
            Next i
            Call bb.init
            Call bb.combine(A1stServeInWithoutAce)                      'Aһ����ACE
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 3) = ba.x
                  .Cells(i + 2, 4) = ba.y
            Next i
            Call bb.init
            Call bb.combine(A2ndServeAce)                               'A����ACE
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 5) = ba.x
                  .Cells(i + 2, 6) = ba.y
            Next i
            Call bb.init
            Call bb.combine(A2ndServeInWithoutAce)                      'A������ACE
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 7) = ba.x
                  .Cells(i + 2, 8) = ba.y
            Next i
            Call bb.init
            Call bb.combine(B1stServeAce)
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 9) = ba.x
                  .Cells(i + 2, 10) = ba.y
            Next i
            Call bb.init
            Call bb.combine(B1stServeInWithoutAce)
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 11) = ba.x
                  .Cells(i + 2, 12) = ba.y
            Next i
            Call bb.init
            Call bb.combine(B2ndServeAce)
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 13) = ba.x
                  .Cells(i + 2, 14) = ba.y
            Next i
            Call bb.init
            Call bb.combine(B2ndServeInWithoutAce)
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 15) = ba.x
                  .Cells(i + 2, 16) = ba.y
            Next i
            Call bb.init
            Call bb.combine(AReturnPoint)
            Call bb.combine(AHitPoint)
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 17) = ba.x
                  .Cells(i + 2, 18) = ba.y
            Next i
            
            Call bb.init
            Call bb.combine(BReturnPoint)
            Call bb.combine(BHitPoint)
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 19) = ba.x
                  .Cells(i + 2, 20) = ba.y
            Next i

      End With
      
Debug_print:
'      Dim ForShow As New vecBall
'      Call ForShow.combine(B1stServeLandingPointMedium)
'      Call ForShow.combine(B2ndServeLandingPointMedium)
'      Debug.Print "��·In"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
'      Call ForShow.combine(B2ndServeLandingPointMediumFault)
'      Call ForShow.combine(B1stServeLandingPointMediumFault)
'      Debug.Print "��·out"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
''      Call ForShow.combine(B1stServeLandingPointMedium)
'      Call ForShow.combine(BHitLandingPointFault)
'      Debug.Print "B����out���"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
''      Call ForShow.init
'      Call ForShow.combine(B1stServeLandingPointMediumFault)
'      Call ForShow.combine(B2ndServeLandingPointMediumFault)
'      Debug.Print "A�ӷ�����In���"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
'      Call ForShow.combine(B1stServeLandingPointOuter)
'      Call ForShow.combine(B2ndServeLandingPointOuter)
'      Debug.Print "���In"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
'      Call ForShow.combine(B1stServeLandingPointOuterFault)
'      Call ForShow.combine(B2ndServeLandingPointOuterFault)
'      Debug.Print "�ڽ�out"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
'      Call ForShow.combine(B1stServeLandingPointOtherFault)
'      Call ForShow.combine(B2ndServeLandingPointOtherFault)
'      Debug.Print "����out"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next


      ''''''''''''''''''''''''''''����ϸ������''''''''''''''''''''''''''''''''''
      With Sheets("raceOtherDetails")
            .Range(.Cells(4, 3), .Cells(32, 4)).Clear
            .Range(.Cells(4, 6), .Cells(32, 7)).Clear
            .Range(.Cells(4, 9), .Cells(32, 9)).Clear
            .Columns("C:D").HorizontalAlignment = xlCenter
            .Columns("F:G").HorizontalAlignment = xlCenter
            .Columns("I:J").HorizontalAlignment = xlCenter
            
            .Cells(4, 3) = A1stServePoint.count - A1stServeLet.count                            'A��һ������
            .Cells(5, 3) = A1stServeAce.count                                                   'A��һ��ACE
            .Cells(6, 3) = A1stServeLandingPointInner.count                                     'A��һ��_�ڽ�In
            .Cells(7, 3) = A1stServeLandingPointMedium.count                                    'A��һ��_��·In
            .Cells(8, 3) = A1stServeLandingPointOuter.count                                     'A��һ��_���In
            .Cells(9, 3) = A1stServeLandingPointOtherFault.count                                'A��һ������Out
            .Cells(10, 3) = A1stServeLandingPointInnerFault.count                               'A��һ��_�ڽ�Out
            .Cells(11, 3) = A1stServeLandingPointMediumFault.count                              'A��һ��_��·Out
            .Cells(12, 3) = A1stServeLandingPointOuterFault.count                               'A��һ��_���Out
            .Cells(13, 3) = A1stServeWin.count                                                  'A��һ���÷ֻغ�
            .Cells(14, 3) = A2ndServePoint.count - A2ndServeLet.count                           'A����������
            .Cells(15, 3) = A2ndServeAce.count                                                  'A������ACE
            .Cells(16, 3) = A2ndServeLandingPointInner.count                                    'A������_�ڽ�In
            .Cells(17, 3) = A2ndServeLandingPointMedium.count                                   'A������_��·In
            .Cells(18, 3) = A2ndServeLandingPointOuter.count                                    'A������_���In
            .Cells(19, 3) = A2ndServeLandingPointOtherFault.count                               'A����������Out
            .Cells(20, 3) = A2ndServeLandingPointInnerFault.count                               'A������_�ڽ�Out
            .Cells(21, 3) = A2ndServeLandingPointMediumFault.count                              'A������_��·Out
            .Cells(22, 3) = A2ndServeLandingPointOuterFault.count                               'A������_���Out
            .Cells(23, 3) = A2ndServeWin.count                                                  'A�������÷ֻغ�

            .Cells(25, 3).FormulaR1C1 = "=sum(R4C3,R14C3)"                                      'A����������
            .Cells(26, 3).FormulaR1C1 = "=sum(R6C3,R16C3)"                                      'A�������ڽ�in
            .Cells(27, 3).FormulaR1C1 = "=sum(R7C3,R17C3)"                                      'A��������·in
            .Cells(28, 3).FormulaR1C1 = "=sum(R8C3,R18C3)"                                      'A���������in
            .Cells(29, 3).FormulaR1C1 = "=sum(R9C3:R12C3,R19C3:R22C3)"                          'A������out
            .Cells(30, 3).FormulaR1C1 = "=IfError(R26C3/(R26C3 +R20C3 +R10C3), 0)"              'A���ڽǳɹ���
            .Cells(30, 3).NumberFormatLocal = "0%"
            .Cells(31, 3).FormulaR1C1 = "=IfError(R27C3/(R27C3 +R21C3 +R11C3), 0)"              'A����·�ɹ���
            .Cells(31, 3).NumberFormatLocal = "0%"
            .Cells(32, 3).FormulaR1C1 = "=IfError(R28C3/(R28C3 +R22C3 +R12C3), 0)"              'A����ǳɹ���
            .Cells(32, 3).NumberFormatLocal = "0%"


            .Cells(4, 6) = AReturnPoint.count                                                   'A���ӷ�����
            .Cells(5, 6) = AReturnLandingPointEasy.count                                        'A���ӷ���
            .Cells(6, 6) = AReturnLandingPointNormal.count                                      'A���ӷ���
            .Cells(7, 6) = AReturnLandingPointHard.count                                        'A���ӷ���
            .Cells(8, 6) = AReturnBeingVolleyPoint.count                                        'A���ӷ����ػ�
            .Cells(9, 6).FormulaR1C1 = "=sum(R5C6,R8C6)"                                        'A���ӷ��ɹ�
            .Cells(10, 6) = AReturnForehandTotal.count                                          'A���ӷ�����
            .Cells(11, 6) = AReturnBackhandTotal.count                                          'A���ӷ�����
            .Cells(12, 6) = AReturnForehandIn.count                                             'A���ӷ�����in
            .Cells(13, 6) = AReturnBackhandIn.count                                             'A���ӷ�����in
            .Cells(14, 6) = AReturnLandingPointFault.count                                      'A���ӷ�out
            .Cells(15, 6).FormulaR1C1 = "=IfError(R9C6/R4C6, 0)"                                'A���ӷ��ɹ���
            .Cells(15, 6).NumberFormatLocal = "0%"

            .Cells(18, 6) = ABreakSucceed.count                                                 'A���Ʒ��÷�
            .Cells(19, 6) = ABreakPoint.count                                                   'A���Ʒ���
            .Cells(20, 6) = ANetNeerByWin.count                                                 'A�������÷�
            .Cells(21, 6) = ANetNeerByPoint.count                                               'A����������
            .Cells(22, 6) = A1stServeLet.count                                                  'Aһ��Let
            .Cells(23, 6) = A2ndServeLet.count                                                  'A����Let

            .Cells(25, 6).FormulaR1C1 = "=sum(R5C3��R15C3)"                                     'A��ACE
            .Cells(26, 6).FormulaR1C1 = "=sum(R19C3:R22C3)"                                     'A��˫��
            .Cells(27, 6).FormulaR1C1 = "=IfError(sum(R6C3:R8C3)/R4C3, 0)"                      'A��һ���ɹ���
            .Cells(27, 6).NumberFormatLocal = "0%"
            .Cells(28, 6).FormulaR1C1 = "=IfError(R13C3/sum(R6C3:R8C3), 0)"                     'A��һ��ʤ��
            .Cells(28, 6).NumberFormatLocal = "0%"
            .Cells(29, 6).FormulaR1C1 = "=IfError(R23C3/sum(R16C3:R18C3), 0)"                   'A������ʤ��
            .Cells(29, 6).NumberFormatLocal = "0%"
            .Cells(30, 6).FormulaR1C1 = "=IfError(R18C6/R19C6, 0)"                              'A���Ʒ��÷���
            .Cells(30, 6).NumberFormatLocal = "0%"
            .Cells(31, 6).FormulaR1C1 = "=IfError(R20C6/R21C6, 0)"                              'A����ǰ�÷���
            .Cells(31, 6).NumberFormatLocal = "0%"
            .Cells(32, 6) = AWinner.count                                                       'A����ʤ��
            
            
            .Cells(4, 9) = AHitPoint.count                                                      'A����������
            .Cells(5, 9) = AHitLandingPointEasy.count                                           'A��������
            .Cells(6, 9) = AHitLandingPointNormal.count                                         'A��������
            .Cells(7, 9) = AHitLandingPointHard.count                                           'A��������
            .Cells(8, 9) = AHitBeingVolleyPoint.count                                           'A�����򱻽ػ�
            .Cells(9, 9).FormulaR1C1 = "=sum(R5C9,R8C9)"                                        'A������ɹ�
            .Cells(10, 9) = AHitForehandTotal.count                                             'A����������
            .Cells(11, 9) = AHitBackhandTotal.count                                             'A��������
            .Cells(12, 9) = AHitForehandIn.count                                                'A����������In
            .Cells(13, 9) = AHitBackhandIn.count                                                'A��������In
            .Cells(14, 9) = AHitLandingPointFault.count                                         'A������out
            .Cells(15, 9).FormulaR1C1 = "=IfError(R12C9/R10C9, 0)"                              'A�������ֳɹ���
            .Cells(15, 9).NumberFormatLocal = "0%"
            .Cells(16, 9).FormulaR1C1 = "=IfError(R13C9/R11C9, 0)"                              'A�����ֳɹ���
            .Cells(16, 9).NumberFormatLocal = "0%"
            .Cells(18, 9).FormulaR1C1 = "=sum(R4C9,R4C6)"                                       'A����������(���ӷ�)
            .Cells(19, 9).FormulaR1C1 = "=sum(R12C9:R12C6)"                                     'A����������in(���ӷ�)
            .Cells(20, 9).FormulaR1C1 = "=sum(R13C9:R13C6)"                                     'A��������in(���ӷ�)
            .Cells(21, 9).FormulaR1C1 = "=sum(R14C9:R14C6)"                                     'A������out(���ӷ�)
            
            .Cells(25, 9) = AShortRoundWin.count                                                'A������ʤ���غ�
            .Cells(26, 9) = AMediumRoundWin.count                                               'A������ʤ���غ�
            .Cells(27, 9) = ALongRoundWin.count                                                 'A������ʤ���غ�
            .Cells(28, 9) = ABoutWin.count                                                      'A���ܵ÷ֻغ���
            
            .Cells(4, 4) = B1stServePoint.count - B1stServeLet.count                            'B��һ������
            .Cells(5, 4) = B1stServeAce.count                                                   'B��һ��ACE
            .Cells(6, 4) = B1stServeLandingPointInner.count                                     'B��һ��_�ڽ�In
            .Cells(7, 4) = B1stServeLandingPointMedium.count                                    'B��һ��_��·In
            .Cells(8, 4) = B1stServeLandingPointOuter.count                                     'B��һ��_���In
            .Cells(9, 4) = B1stServeLandingPointOtherFault.count                                'B��һ������Out
            .Cells(10, 4) = B1stServeLandingPointInnerFault.count                               'B��һ��_�ڽ�Out
            .Cells(11, 4) = B1stServeLandingPointMediumFault.count                              'B��һ��_��·Out
            .Cells(12, 4) = B1stServeLandingPointOuterFault.count                               'B��һ��_���Out
            .Cells(13, 4) = B1stServeWin.count                                                  'B��һ���÷ֻغ�
            .Cells(14, 4) = B2ndServePoint.count - B2ndServeLet.count                           'B����������
            .Cells(15, 4) = B2ndServeAce.count                                                  'B������ACE
            .Cells(16, 4) = B2ndServeLandingPointInner.count                                    'B������_�ڽ�In
            .Cells(17, 4) = B2ndServeLandingPointMedium.count                                   'B������_��·In
            .Cells(18, 4) = B2ndServeLandingPointOuter.count                                    'B������_���In
            .Cells(19, 4) = B2ndServeLandingPointOtherFault.count                               'B����������Out
            .Cells(20, 4) = B2ndServeLandingPointInnerFault.count                               'B������_�ڽ�Out
            .Cells(21, 4) = B2ndServeLandingPointMediumFault.count                              'B������_��·Out
            .Cells(22, 4) = B2ndServeLandingPointOuterFault.count                               'B������_���Out
            .Cells(23, 4) = B2ndServeWin.count                                                  'B�������÷ֻغ�

            .Cells(25, 4).FormulaR1C1 = "=sum(R4C4,R14C4)"                                      'B����������
            .Cells(26, 4).FormulaR1C1 = "=sum(R6C4,R16C4)"                                      'B�������ڽ�in
            .Cells(27, 4).FormulaR1C1 = "=sum(R7C4,R17C4)"                                      'B��������·in
            .Cells(28, 4).FormulaR1C1 = "=sum(R8C4,R18C4)"                                      'B���������in
            .Cells(29, 4).FormulaR1C1 = "=sum(R9C4:R12C4,R19C4:R22C4)"                          'B������out
            .Cells(30, 4).FormulaR1C1 = "=IfError(R26C4/(R26C4 +R20C4 +R10C4), 0)"              'B���ڽǳɹ���
            .Cells(30, 4).NumberFormatLocal = "0%"
            .Cells(31, 4).FormulaR1C1 = "=IfError(R27C4/(R27C4 +R21C4 +R11C4), 0)"              'B����·�ɹ���
            .Cells(31, 4).NumberFormatLocal = "0%"
            .Cells(32, 4).FormulaR1C1 = "=IfError(R28C4/(R28C4 +R22C4 +R12C4), 0)"              'B����ǳɹ���
            .Cells(32, 4).NumberFormatLocal = "0%"


            .Cells(4, 7) = BReturnPoint.count                                                   'B���ӷ�����
            .Cells(5, 7) = BReturnLandingPointEasy.count                                        'B���ӷ���
            .Cells(6, 7) = BReturnLandingPointNormal.count                                      'B���ӷ���
            .Cells(7, 7) = BReturnLandingPointHard.count                                        'B���ӷ���
            .Cells(8, 7) = BReturnBeingVolleyPoint.count                                        'B���ӷ����ػ�
            .Cells(9, 7).FormulaR1C1 = "=sum(R5C7,R8C7)"                                        'B���ӷ��ɹ�
            .Cells(10, 7) = BReturnForehandTotal.count                                          'B���ӷ�����
            .Cells(11, 7) = BReturnBackhandTotal.count                                          'B���ӷ�����
            .Cells(12, 7) = BReturnForehandIn.count                                             'B���ӷ�����in
            .Cells(13, 7) = BReturnBackhandIn.count                                             'B���ӷ�����in
            .Cells(14, 7) = BReturnLandingPointFault.count                                      'B���ӷ�out
            .Cells(15, 7).FormulaR1C1 = "=IfError(R9C7/R4C7, 0)"                                'B���ӷ��ɹ���
            .Cells(15, 7).NumberFormatLocal = "0%"

            .Cells(18, 7) = BBreakSucceed.count                                                 'B���Ʒ��÷�
            .Cells(19, 7) = BBreakPoint.count                                                   'B���Ʒ���
            .Cells(20, 7) = BNetNeerByWin.count                                                'B�������÷�
            .Cells(21, 7) = BNetNeerByPoint.count                                                  'B����������
            .Cells(22, 7) = B1stServeLet.count                                                  'Bһ��Let
            .Cells(23, 7) = B2ndServeLet.count                                                  'B����Let

            .Cells(25, 7).FormulaR1C1 = "=sum(R5C4,R15C4)"                                      'B��ACE
            .Cells(26, 7).FormulaR1C1 = "=sum(R19C4:R22C4)"                                     'B��˫��
            .Cells(27, 7).FormulaR1C1 = "=IfError(sum(R6C4:R8C4)/R4C4, 0)"                      'B��һ���ɹ���
            .Cells(27, 7).NumberFormatLocal = "0%"
            .Cells(28, 7).FormulaR1C1 = "=IfError(R13C4/sum(R6C4:R8C4), 0)"                     'B��һ��ʤ��
            .Cells(28, 7).NumberFormatLocal = "0%"
            .Cells(29, 7).FormulaR1C1 = "=IfError(R23C4/sum(R16C4:R18C4), 0)"                   'B������ʤ��
            .Cells(29, 7).NumberFormatLocal = "0%"
            .Cells(30, 7).FormulaR1C1 = "=IfError(R18C7/R19C7, 0)"                              'B���Ʒ��÷���
            .Cells(30, 7).NumberFormatLocal = "0%"
            .Cells(31, 7).FormulaR1C1 = "=IfError(R20C7/R21C7, 0)"                              'B����ǰ�÷���
            .Cells(31, 7).NumberFormatLocal = "0%"
            .Cells(32, 7) = BWinner.count                                                       'B����ʤ��
            
            
            .Cells(4, 10) = BHitPoint.count                                                      'B����������
            .Cells(5, 10) = BHitLandingPointEasy.count                                           'B��������
            .Cells(6, 10) = BHitLandingPointNormal.count                                         'B��������
            .Cells(7, 10) = BHitLandingPointHard.count                                           'B��������
            .Cells(8, 10) = BHitBeingVolleyPoint.count                                           'B�����򱻽ػ�
            .Cells(9, 10).FormulaR1C1 = "=sum(R5C10,R8C10)"                                      'B������ɹ�
            .Cells(10, 10) = BHitForehandTotal.count                                             'B����������
            .Cells(11, 10) = BHitBackhandTotal.count                                             'B��������
            .Cells(12, 10) = BHitForehandIn.count                                                'B����������In
            .Cells(13, 10) = BHitBackhandIn.count                                                'B��������In
            .Cells(14, 10) = BHitLandingPointFault.count                                         'B������out
            .Cells(15, 10).FormulaR1C1 = "=IfError(R12C10/R10C10, 0)"                            'B�������ֳɹ���
            .Cells(15, 10).NumberFormatLocal = "0%"
            .Cells(16, 10).FormulaR1C1 = "=IfError(R13C10/R11C10, 0)"                            'B�����ֳɹ���
            .Cells(16, 10).NumberFormatLocal = "0%"
            .Cells(18, 10).FormulaR1C1 = "=sum(R4C10,R4C7)"                                      'A����������(���ӷ�)
            .Cells(19, 10).FormulaR1C1 = "=sum(R12C10,R12C7)"                                    'A����������in(���ӷ�)
            .Cells(20, 10).FormulaR1C1 = "=sum(R13C10,R13C7)"                                    'A��������in(���ӷ�)
            .Cells(21, 10).FormulaR1C1 = "=sum(R14C10,R14C7)"                                    'A������out(���ӷ�)
            
            .Cells(25, 10) = BShortRoundWin.count                                                'B������ʤ���غ�
            .Cells(26, 10) = BMediumRoundWin.count                                               'B������ʤ���غ�
            .Cells(27, 10) = BLongRoundWin.count                                                 'B������ʤ���غ�
            .Cells(28, 10) = BBoutWin.count                                                      'B���ܵ÷ֻغ���
            
      End With
      
End Sub

Sub breakPointGrabber(j%, gameMode%, isWinnerServe As Boolean, ba As ball, _
                  WinnerBout%, WinnerGame%, WinnerBreakPoint As vecBall, WinnerBreakSucceed As vecBall, _
                  serveFlag%, Winner1stServeWin As vecBall, Winner2ndServeWin As vecBall, WinnerBoutWin As vecBall, _
                  LoserBout%, LoserGame%)
      If gameMode = 4 Or gameMode = 6 Then                                                      '������ģʽ
            If WinnerGame = gameMode And WinnerGame = LoserGame Then                            '�������߾�
                  If WinnerBout >= 7 And WinnerBout >= LoserBout + 2 Then
                        LoserBout = 0
                        LoserGame = 0
                        WinnerBout = 0
                        WinnerGame = 0
                  End If
            Else                                                                                '������ͨ��
                  If WinnerBout >= 4 And WinnerBout >= LoserBout + 2 Then
                        WinnerGame = WinnerGame + 1
                        LoserBout = 0
                        WinnerBout = 0
                        If Not isWinnerServe Then
                              Call WinnerBreakSucceed.push_back(ba)
                              Sheets("main").Cells(j, 11).Value = "BreakPointWin!"
                        End If
                        If WinnerGame >= gameMode And WinnerGame >= LoserGame + 2 Then
                              WinnerGame = 0
                              LoserGame = 0
                        End If
                  ElseIf WinnerBout >= 3 And WinnerBout > LoserBout Then
                        If Not isWinnerServe Then
                              Call WinnerBreakPoint.push_back(ba)
                              Sheets("main").Cells(j, 11).Value = "BreakPoint"
                        End If
                  End If
            End If
      End If
      If isWinnerServe Then
            If Abs(serveFlag) = 1 Then
                  Call Winner1stServeWin.push_back(ba)
            Else
                  Call Winner2ndServeWin.push_back(ba)
            End If
      End If
      Call WinnerBoutWin.push_back(ba)
End Sub




