Attribute VB_Name = "模块4"
Option Explicit

Const VOLLEY As Integer = 100
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''比赛统计部分'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub raceAnalysis()
      Dim aim As Range
      Dim i%, j%, k%, m%, FOREBACK%
      
      Dim ba As New ball
      
      Dim latestLandingIn As New ball           '上一个球落点，用于记录当前球并不是落点但是回合结束(如ACE或者致胜分)时的数据信息
      Dim latestHit As New ball                 '上一个主动动作击球点(发球或者击球)
      Dim intervalHit As New ball               '上上个主动动作击球点(发球或者击球)，用于记录在统计接发球或者击球难度时对方击球位置参数
      
      Dim Flag%, latestHitStat%, serveFlag%, errorCount%, roundCount%, latestHitCommit$
      'flag用于标识有无ACE或者制胜分的可能，当前为落地In时flag=1,为主动动作(发球，击球)时flag=0
      'latestHitStat用于传递本次击球方(+ or -)和是击球事件还是发球事件(1 or 2)给下面的逻辑
      'serveFlag用来标识本回合发球方(+ or -)以及一发(1)二发(2)，发球置时改变
      'errorCount用来记录出了多少个错
      'roundCount用来记录本回合一共打了多少个来回，以接发球方(A+ B-)的击球数为准
      'latestHitCommit记录最近一次击球描述，用于区分是普通回击("hitBack")还是接发球("return")
      Dim A_bout%, B_bout%, A_game%, B_game%, gameMode%
      
      Dim A1stServePoint As New vecBall                           'A_一发发球点
      Dim A1stServeLandingPointOtherFault As New vecBall          'A_一发失误其他落点
      Dim A1stServeLandingPointInnerFault As New vecBall          'A_一发_内角失误落点
      Dim A1stServeLandingPointMediumFault As New vecBall         'A_一发_中路失误落点
      Dim A1stServeLandingPointOuterFault As New vecBall          'A_一发_外角失误落点
      Dim A1stServeLandingPointInner As New vecBall               'A_一发落点_内
      Dim A1stServeLandingPointMedium As New vecBall              'A_一发落点_中
      Dim A1stServeLandingPointOuter As New vecBall               'A_一发落点_外
      Dim A1stServeLet As New vecBall                             'A_一发网球
      Dim A2ndServeLet As New vecBall                             'A_二发网球
      Dim A2ndServePoint As New vecBall                           'A_二发发球点
      Dim A2ndServeLandingPointOtherFault As New vecBall          'A_双误其他落点
      Dim A2ndServeLandingPointInnerFault As New vecBall          'A_二发_内角失误落点
      Dim A2ndServeLandingPointMediumFault As New vecBall         'A_二发_中路失误落点
      Dim A2ndServeLandingPointOuterFault As New vecBall          'A_二发_外角失误落点
      Dim A2ndServeLandingPointInner As New vecBall               'A_二发落点_内
      Dim A2ndServeLandingPointMedium As New vecBall              'A_二发落点_中
      Dim A2ndServeLandingPointOuter As New vecBall               'A_二发落点_外
      Dim AReturnPoint As New vecBall                             'A_接发球击球点
      Dim AReturnLandingPointEasy As New vecBall                  'A_接发球落点_易
      Dim AReturnLandingPointNormal As New vecBall                'A_接发球落点_中
      Dim AReturnLandingPointHard As New vecBall                  'A_接发球落点_难
      Dim AReturnLandingPointFault As New vecBall                 'A_接发球失误落点
      Dim AReturnBeingVolleyPoint As New vecBall                  'A_接发球被对方截击
      Dim AHitBeingVolleyPoint As New vecBall                     'A_击球被对方截击
      Dim AHitPoint As New vecBall                                'A_击球点
      Dim AHitLandingPointEasy As New vecBall                     'A_击球落点_易
      Dim AHitLandingPointNormal As New vecBall                   'A_击球落点_中
      Dim AHitLandingPointHard As New vecBall                     'A_击球落点_难
      Dim AHitLandingPointFault As New vecBall                    'A_击球失误落点
      Dim ANetNeerByPoint As New vecBall                             'A_上网击球点
      Dim ANetNeerByWin As New vecBall                           'A_网前得分
'      Dim AAce As New vecBall                                     'A_ace
      Dim A1stServeAce As New vecBall                             'A_一发ace
      Dim A2ndServeAce As New vecBall                             'A_二发ace
      Dim AWinner As New vecBall                                  'A_制胜分
      Dim ABreakPoint As New vecBall                              'A_破发点
      Dim ABreakSucceed As New vecBall                            'A_破发得分
      Dim AShortRoundWin As New vecBall                           'A_短拍胜利回合数
      Dim AMediumRoundWin As New vecBall                          'A_中拍胜利回合数
      Dim ALongRoundWin As New vecBall                            'A_长拍胜利回合数
      Dim A1stServeWin As New vecBall                             'A_一发得分回合数
      Dim A2ndServeWin As New vecBall                             'A_二发得分回合数
      Dim ABoutWin As New vecBall                                 'A_总得分数
      
      Dim A1stServeInWithoutAce As New vecBall                    'A一发非ACE球落点       //为了落点统计加入的，无其他意义
      Dim A2ndServeInWithoutAce As New vecBall                    'A二发非ACE球落点
      Dim B1stServeInWithoutAce As New vecBall                    'B一发非ACE球落点
      Dim B2ndServeInWithoutAce As New vecBall                    'B二发非ACE球落点
      
      Dim B1stServePoint As New vecBall                           'B_一发发球点
      Dim B1stServeLandingPointOtherFault As New vecBall          'B_一发失误其他落点
      Dim B1stServeLandingPointInnerFault As New vecBall          'B_一发_内角失误落点
      Dim B1stServeLandingPointMediumFault As New vecBall         'B_一发_中路失误落点
      Dim B1stServeLandingPointOuterFault As New vecBall          'B_一发_外角失误落点
      Dim B1stServeLandingPointInner As New vecBall               'B_一发落点_内
      Dim B1stServeLandingPointMedium As New vecBall              'B_一发落点_中
      Dim B1stServeLandingPointOuter As New vecBall               'B_一发落点_外
      Dim B1stServeLet As New vecBall                             'B_一发网球
      Dim B2ndServeLet As New vecBall                             'B_二发网球
      Dim B2ndServePoint As New vecBall                           'B_二发发球点
      Dim B2ndServeLandingPointOtherFault As New vecBall          'B_双误其他落点
      Dim B2ndServeLandingPointInnerFault As New vecBall          'B_二发_内角失误落点
      Dim B2ndServeLandingPointMediumFault As New vecBall         'B_二发_中路失误落点
      Dim B2ndServeLandingPointOuterFault As New vecBall          'B_二发_外角失误落点
      Dim B2ndServeLandingPointInner As New vecBall               'B_二发落点_内
      Dim B2ndServeLandingPointMedium As New vecBall              'B_二发落点_中
      Dim B2ndServeLandingPointOuter As New vecBall               'B_二发落点_外
      Dim BReturnPoint As New vecBall                             'B_接发球击球点
      Dim BReturnLandingPointEasy As New vecBall                  'B_接发球落点_易
      Dim BReturnLandingPointNormal As New vecBall                'B_接发球落点_中
      Dim BReturnLandingPointHard As New vecBall                  'B_接发球落点_难
      Dim BReturnLandingPointFault As New vecBall                 'B_接发球失误落点
      Dim BReturnBeingVolleyPoint As New vecBall                  'B_接发球被对方截击
      Dim BHitBeingVolleyPoint As New vecBall                     'B_击球被对方截击
      Dim BHitPoint As New vecBall                                'B_击球点
      Dim BHitLandingPointEasy As New vecBall                     'B_击球落点_易
      Dim BHitLandingPointNormal As New vecBall                   'B_击球落点_中
      Dim BHitLandingPointHard As New vecBall                     'B_击球落点_难
      Dim BHitLandingPointFault As New vecBall                    'B_击球失误落点
      Dim BNetNeerByPoint As New vecBall                             'B_上网击球点
      Dim BNetNeerByWin As New vecBall                           'B_网前得分
'      Dim BAce As New vecBall                                     'B_ace
      Dim B1stServeAce As New vecBall                             'B_一发ace
      Dim B2ndServeAce As New vecBall                             'B_二发ace
      Dim BWinner As New vecBall                                  'B_制胜分
      Dim BBreakPoint As New vecBall                              'B_破发点
      Dim BBreakSucceed As New vecBall                            'B_破发得分
      Dim BShortRoundWin As New vecBall                           'B_短拍胜利回合数
      Dim BMediumRoundWin As New vecBall                          'B_中拍胜利回合数
      Dim BLongRoundWin As New vecBall                            'B_长拍胜利回合数
      Dim B1stServeWin As New vecBall                             'B_一发得分回合数
      Dim B2ndServeWin As New vecBall                             'B_二发得分回合数
      Dim BBoutWin As New vecBall                                 'B_总得分数
      
      Dim inner As New multiZone                                  '发球_内角区域
      Dim medium As New multiZone                                 '发球_中路区域
      Dim outer As New multiZone                                  '发球_外角区域
      
      Dim innerFaultAntiB1 As New multiZone                       '在B1区发球时判定为发球_内角失误的区域
      Dim innerFaultAntiB2 As New multiZone                       '在B2区发球时判定为发球_内角失误的区域
      Dim innerFaultAntiA1 As New multiZone                       '在A1区发球时判定为发球_内角失误的区域
      Dim innerFaultAntiA2 As New multiZone                       '在A2区发球时判定为发球_内角失误的区域
      Dim mediumFaultAntiB1 As New multiZone                      '在B1区发球时判定为发球_中路失误的区域
      Dim mediumFaultAntiB2 As New multiZone                      '在B2区发球时判定为发球_中路失误的区域
      Dim mediumFaultAntiA1 As New multiZone                      '在A1区发球时判定为发球_中路失误的区域
      Dim mediumFaultAntiA2 As New multiZone                      '在A2区发球时判定为发球_中路失误的区域
      Dim outerFaultAntiB1 As New multiZone                       '在B1区发球时判定为发球_外角失误的区域
      Dim outerFaultAntiB2 As New multiZone                       '在B2区发球时判定为发球_外角失误的区域
      Dim outerFaultAntiA1 As New multiZone                       '在A1区发球时判定为发球_外角失误的区域
      Dim outerFaultAntiA2 As New multiZone                       '在A2区发球时判定为发球_外角失误的区域
      
      Dim normalAntiB1 As New multiZone                           '对手位于B1区击球时本方回击落点目标区域 _中
      Dim hardAntiB1 As New multiZone                             '对手位于B1区击球时本方回击落点目标区域 _难
      Dim normalAntiB2 As New multiZone                           '对手位于B2区击球时本方回击落点目标区域 _中
      Dim hardAntiB2 As New multiZone                             '对手位于B2区击球时本方回击落点目标区域 _难
      Dim normalAntiA1 As New multiZone                           '对手位于A1区击球时本方回击落点目标区域 _中
      Dim hardAntiA1 As New multiZone                             '对手位于A1区击球时本方回击落点目标区域 _难
      Dim normalAntiA2 As New multiZone                           '对手位于A2区击球时本方回击落点目标区域 _中
      Dim hardAntiA2 As New multiZone                             '对手位于A2区击球时本方回击落点目标区域 _难
      '因为发球难度区域在四个象限彼此不交叉所以只需要设置三个复合区域，但是击球难度区域在四个象限交叉，所以要分别设置
      
      Dim tmpZone As New zone
      Dim tempMultiZone As New multiZone
      
      Range(Sheets("main").Columns(8), Sheets("main").Columns(11)).Clear
      If Sheets("main").Range("N3").Value = 1 Then                'cm坐标
            FOREBACK = 274
            Call tmpZone.init(0, 50, 640, 0)                      '发球_内角在A1区的范围
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorX                                  '发球_内角在A2区的范围
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorY                                  '发球_内角在B1区的范围
            Call inner.push_back(tmpZone)
            Call tmpZone.mirrorX                                  '发球_内角在B2区的范围
            Call inner.push_back(tmpZone)
            Call tmpZone.init(0, 361, 640, 50)
            Call medium.push_back(tmpZone)                        '发球中在A1区的范围
            Call tmpZone.mirrorX
            Call medium.push_back(tmpZone)
            Call tmpZone.mirrorY
            Call medium.push_back(tmpZone)
            Call tmpZone.mirrorX
            Call medium.push_back(tmpZone)
            Call tmpZone.init(0, 411, 640, 361)
            Call outer.push_back(tmpZone)                         '发球外在A1区的范围
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
            Call innerFaultAntiB1.clone(tempMultiZone)            '在B1区发球时判定为发球_内角失误的区域
            Call tempMultiZone.mirrorX
            Call innerFaultAntiB2.clone(tempMultiZone)            '在B2区发球时判定为发球_内角失误的区域
            Call tempMultiZone.mirrorY
            Call innerFaultAntiA1.clone(tempMultiZone)            '在A1区发球时判定为发球_内角失误的区域
            Call tempMultiZone.mirrorX
            Call innerFaultAntiA2.clone(tempMultiZone)            '在A2区发球时判定为发球_内角失误的区域
            Call tmpZone.init(640, 361, 1588, 50)
            Call mediumFaultAntiB1.push_back(tmpZone)             '在B1区发球时判定为发球_中路失误的区域
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
            Call outerFaultAntiB1.clone(tempMultiZone)            '在B1区发球时判定为发球_外角失误的区域
            Call tempMultiZone.mirrorX
            Call outerFaultAntiB2.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call outerFaultAntiA1.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call outerFaultAntiA2.combine(tempMultiZone)
            
            If Sheets("main").Range("M3").Value = 1 Then          '单打
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-1188, 411, 0, 229)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, 411, -1006, -411)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, -361, 0, -411)
                  Call tempMultiZone.push_back(tmpZone)
                  Call hardAntiB1.clone(tempMultiZone)        '对手在B1区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorX
                  Call hardAntiB2.clone(tempMultiZone)        '对手在B2区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorY
                  Call hardAntiA1.clone(tempMultiZone)        '对手在A1区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorX
                  Call hardAntiA2.clone(tempMultiZone)        '对手在A2区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-1188, 411, 0, 229)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, 411, -789, -411)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, -361, 0, -411)
                  Call tempMultiZone.push_back(tmpZone)
                  Call normalAntiB1.clone(tempMultiZone)      '对手在B1区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorX
                  Call normalAntiB2.clone(tempMultiZone)      '对手在B2区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorY
                  Call normalAntiA1.clone(tempMultiZone)      '对手在A1区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorX
                  Call normalAntiA2.clone(tempMultiZone)      '对手在A2区击球时本方回球难度为中的目标区域
            Else      '双打
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-1188, 548, 0, 307)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, 548, -1006, -483)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, -483, 0, -548)
                  Call tempMultiZone.push_back(tmpZone)
                  Call hardAntiB1.clone(tempMultiZone)        '对手在B1区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorX
                  Call hardAntiB2.clone(tempMultiZone)        '对手在B2区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorY
                  Call hardAntiA1.clone(tempMultiZone)        '对手在A1区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorX
                  Call hardAntiA2.clone(tempMultiZone)        '对手在A2区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-1188, 548, 0, 307)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, 548, -789, -483)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-1188, -483, 0, -548)
                  Call tempMultiZone.push_back(tmpZone)
                  Call normalAntiB1.clone(tempMultiZone)      '对手在B1区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorX
                  Call normalAntiB2.clone(tempMultiZone)      '对手在B2区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorY
                  Call normalAntiA1.clone(tempMultiZone)      '对手在A1区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorX
                  Call normalAntiA2.clone(tempMultiZone)      '对手在A2区击球时本方回球难度为中的目标区域
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
            Call innerFaultAntiB1.clone(tempMultiZone)            '在B1区发球时判定为发球_内角失误的区域
            Call tempMultiZone.mirrorX
            Call innerFaultAntiB2.clone(tempMultiZone)            '在B2区发球时判定为发球_内角失误的区域
            Call tempMultiZone.mirrorY
            Call innerFaultAntiA1.clone(tempMultiZone)            '在A1区发球时判定为发球_内角失误的区域
            Call tempMultiZone.mirrorX
            Call innerFaultAntiA2.clone(tempMultiZone)            '在A2区发球时判定为发球_内角失误的区域
            Call tmpZone.init(6401, 3615, 15887, 500)
            Call mediumFaultAntiB1.push_back(tmpZone)             '在B1区发球时判定为发球_中路失误的区域
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
            Call outerFaultAntiB1.clone(tempMultiZone)            '在B1区发球时判定为发球_外角失误的区域
            Call tempMultiZone.mirrorX
            Call outerFaultAntiB2.combine(tempMultiZone)
            Call tempMultiZone.mirrorY
            Call outerFaultAntiA1.combine(tempMultiZone)
            Call tempMultiZone.mirrorX
            Call outerFaultAntiA2.combine(tempMultiZone)
            
            If Sheets("main").Range("M3").Value = 1 Then          '单打
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-11887, 4115, 0, 2286)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, 4115, -10059, -4115)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, -3615, 0, -4115)
                  Call tempMultiZone.push_back(tmpZone)
                  Call hardAntiB1.clone(tempMultiZone)        '对手在B1区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorX
                  Call hardAntiB2.clone(tempMultiZone)        '对手在B2区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorY
                  Call hardAntiA1.clone(tempMultiZone)        '对手在A1区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorX
                  Call hardAntiA2.clone(tempMultiZone)        '对手在A2区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-11887, 4115, 0, 2286)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, 4115, -7890, -4115)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, -3615, 0, -4115)
                  Call tempMultiZone.push_back(tmpZone)
                  Call normalAntiB1.clone(tempMultiZone)      '对手在B1区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorX
                  Call normalAntiB2.clone(tempMultiZone)      '对手在B2区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorY
                  Call normalAntiA1.clone(tempMultiZone)      '对手在A1区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorX
                  Call normalAntiA2.clone(tempMultiZone)      '对手在A2区击球时本方回球难度为中的目标区域
            Else      '双打
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-11887, 5487, 0, 3073)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, 5487, -10059, -4829)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, -4829, 0, -5487)
                  Call tempMultiZone.push_back(tmpZone)
                  Call hardAntiB1.clone(tempMultiZone)        '对手在B1区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorX
                  Call hardAntiB2.clone(tempMultiZone)        '对手在B2区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorY
                  Call hardAntiA1.clone(tempMultiZone)        '对手在A1区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.mirrorX
                  Call hardAntiA2.clone(tempMultiZone)        '对手在A2区击球时本方回球难度为难的目标区域
                  Call tempMultiZone.Clear
                  Call tmpZone.init(-11887, 5487, 0, 3073)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, 5487, -789, -4829)
                  Call tempMultiZone.push_back(tmpZone)
                  Call tmpZone.init(-11887, -4829, 0, -5487)
                  Call tempMultiZone.push_back(tmpZone)
                  Call normalAntiB1.clone(tempMultiZone)      '对手在B1区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorX
                  Call normalAntiB2.clone(tempMultiZone)      '对手在B2区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorY
                  Call normalAntiA1.clone(tempMultiZone)      '对手在A1区击球时本方回球难度为中的目标区域
                  Call tempMultiZone.mirrorX
                  Call normalAntiA2.clone(tempMultiZone)      '对手在A2区击球时本方回球难度为中的目标区域
            End If
      Else
            MsgBox "cm/mm数值设置不对"
            Exit Sub
      End If
      
      latestHitStat = 0                   '标识上主动动作击球方以及是击球还是发球
      serveFlag = 0                       '标识一发二发
      Flag = 0                            '标识有没有制胜分或者ACE球的可能
      A_bout = 0
      B_bout = 0
      A_game = 0
      B_game = 0
      roundCount = 0
      gameMode = IIf(Sheets("main").Range("Q15").Value = "", 6, Sheets("main").Range("Q15").Value)    '比赛模式默认六局抢七
      
      
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
      '上一个击球的正反手性质：正手true;反手击球false
      
      
      With Sheets("main")
            .Columns("h").Clear
            Set aim = .Range(.Cells(1, 7), .Cells(9999, 7).End(xlUp))
            For i = 1 To aim.Cells.count
                  If aim.Cells(i) <> "" Then
                        j = aim.Cells(i).Row
                        Call ba.init(.Cells(j, 1), .Cells(j, 2), .Cells(j, 3), .Cells(j, 4), .Cells(j, 5))
                        If .Cells(j, 6) Like "Error*" Then                          '错误
                              errorCount = errorCount + 1
                        ElseIf .Cells(j, 6) = "firstServe" Then                     '一发
                              If .Cells(j, 1) * .Cells(j, 7) > 0 Then   'A发球
                                    serveFlag = 1
                                    latestHitStat = 2
                                    Call A1stServePoint.push_back(ba)
                                    .Cells(j, 8) = "A1stServePoint"
                              Else                                      'B发球
                                    serveFlag = -1
                                    latestHitStat = -2
                                    Call B1stServePoint.push_back(ba)
                                    .Cells(j, 8) = "B1stServePoint"
                              End If
                              Flag = 0
                              Call latestHit.clone(ba)
                              latestHitCommit = .Cells(j, 6)
                        ElseIf .Cells(j, 6) = "secondServe" Then                    '二发
                              If .Cells(j, 1) * .Cells(j, 7) > 0 Then   'A发球
                                    serveFlag = 2
                                    latestHitStat = 2
                                    Call A2ndServePoint.push_back(ba)
                                    .Cells(j, 8) = "A2ndServePoint"
                              Else                                      'B发球
                                    serveFlag = -2
                                    latestHitStat = -2
                                    Call B2ndServePoint.push_back(ba)
                                    .Cells(j, 8) = "B2ndServePoint"
                              End If
                              Flag = 0
                              Call latestHit.clone(ba)
                              latestHitCommit = .Cells(j, 6)
                        ElseIf .Cells(j, 6) Like "let*" Then                        '网球
                              If serveFlag = 2 Then         'A二发
                                    Call A2ndServeLet.push_back(ba)
                              ElseIf serveFlag = -2 Then    'B二发
                                    Call B2ndServeLet.push_back(ba)
                              ElseIf serveFlag = 1 Then     'A一发
                                    Call A1stServeLet.push_back(ba)
                              ElseIf serveFlag = -1 Then     'B一发
                                    Call B1stServeLet.push_back(ba)
                              Else
                                    .Cells(j, 8) = "9999999999999999999999999999999999999999999"
                              End If
                        ElseIf .Cells(j, 6) = "firstServeIn" Then                   '一发成功
                              If .Cells(j, 1) * .Cells(j, 7) < 0 Then   '当前落点位于B选手所在方位
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
                              Else                                      '当前落点位于A选手所在方位
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
                        ElseIf .Cells(j, 6) = "secondServeIn" Then                  '二发成功
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
                        ElseIf .Cells(j, 6) = "return" Then                         '接发球
                              Flag = 0
                              Call intervalHit.clone(latestHit)
                              Call latestHit.clone(ba)
                              latestHitCommit = .Cells(j, 6)
                              If .Cells(j, 1) * .Cells(j, 7) > 0 Then   'A方接发球
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
                              Else                                      'B方接发球
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
                        ElseIf .Cells(j, 6) = "hitBack" Then                        '回球
                              '如果当前击球为截击，视上一个击球为IN
                              If Flag = 0 Then
                                    If latestHitCommit = "return" Then   '接发被截击
                                          If .Cells(j, 1) * .Cells(j, 7) > 0 Then   '当前为A方击球
                                                Call BReturnBeingVolleyPoint.push_back(latestHit)
                                                If isLastHitForehand Then
                                                      Call BReturnForehandIn.push_back(latestHit)
                                                Else
                                                      Call BReturnBackhandIn.push_back(latestHit)
                                                End If
                                          Else                                      '当前为B方击球
                                                Call AReturnBeingVolleyPoint.push_back(latestHit)
                                                If isLastHitForehand Then
                                                      Call AReturnForehandIn.push_back(latestHit)
                                                Else
                                                      Call AReturnBackhandIn.push_back(latestHit)
                                                End If
                                          End If
                                    ElseIf latestHitCommit = "hitBack" Then   '击球被截击
                                          If .Cells(j, 1) * .Cells(j, 7) > 0 Then   '当前为A方击球
                                                Call BHitBeingVolleyPoint.push_back(latestHit)
                                                If isLastHitForehand Then
                                                      Call BHitForehandIn.push_back(latestHit)
                                                Else
                                                      Call BHitBackhandIn.push_back(latestHit)
                                                End If
                                          Else                                      '当前为B方击球
                                                Call AHitBeingVolleyPoint.push_back(latestHit)
                                                If isLastHitForehand Then
                                                      Call AHitForehandIn.push_back(latestHit)
                                                Else
                                                      Call AHitBackhandIn.push_back(latestHit)
                                                End If
                                          End If
                                    End If
                              End If
                              '开始处理当前击球数据
                              If .Cells(j, 1) * .Cells(j, 7) > 0 Then   '当前为A方击球
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
                              Else                                      '当前为B方击球
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
                        ElseIf .Cells(j, 6) = "in" Then                             '入区
                              Call latestLandingIn.clone(ba)
                              Flag = 1
                              If .Cells(j, 1) * .Cells(j, 7) < 0 Then     '当前落点在B球手所在区域
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
                              Else                    '当前落点在A方所在区域
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
                        ElseIf .Cells(j, 6) = "fault,waitingForSecondServe" Or .Cells(j, 6) = "faultGuess,waitingForSecondServe" Then      '一发失败
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
                        ElseIf .Cells(j, 6) Like "*boutEnd" Then                                            '回合结束
                              If Flag = 1 Then                          '上一次击球成功落地
                                    If latestHitStat = 2 Then           '最近的一次击球是A方发球 ace
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
                                    ElseIf latestHitStat = -2 Then      '最近的一次击球是B方发球 ace
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
                                    ElseIf latestHitStat = 1 Then       '最近的一次击球是A方击球 winner
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
                                    ElseIf latestHitStat = -1 Then      '最近的一次击球是B方击球 winner
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
                              Else                                      '上一次击球没有成功落地
                                    If latestHitStat = -2 Then          '最近的一次击球是B方发球,doubleFault
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
                                    ElseIf latestHitStat = 2 Then       '最近的一次击球是A方发球,doubleFault
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
                                    ElseIf latestHitStat = -1 Then      '最近的一次击球是B方击球 fault
                                          A_bout = A_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) > 0, ba, A_bout, A_game, ABreakPoint, ABreakSucceed, _
                                          serveFlag, A1stServeWin, A2ndServeWin, ABoutWin, B_bout, B_game)
                                          If latestHitCommit = "return" Then        '上一次击球为接发球
                                                Call BReturnLandingPointFault.push_back(latestHit)
                                                .Cells(j, 8) = "BReturnLandingPointFault"
                                          Else                                      '上一次击球为普通击球
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
                                    ElseIf latestHitStat = 1 Then       '最近的一次击球是A方击球 fault
                                          B_bout = B_bout + 1
                                          Call breakPointGrabber(j, gameMode, Sgn(serveFlag) < 0, ba, B_bout, B_game, BBreakPoint, BBreakSucceed, _
                                          serveFlag, B1stServeWin, B2ndServeWin, BBoutWin, A_bout, A_game)
                                          If latestHitCommit = "return" Then        '上一次击球为接发球
                                                Call AReturnLandingPointFault.push_back(latestHit)
                                                .Cells(j, 8) = "AReturnLandingPointFault"
                                          Else                                      '上一次击球为普通击球
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
      
      '''''''''''''''''''''''''''''''''''''发球与击球落点呈现''''''''''''''''''''''''''''''''''''''''
      With Sheets("raceCoordinates")
            .Columns("A:T").Clear
            .Cells(1, 1) = "A一发ACE"
            .Cells(1, 3) = "A一发非ACE"
            .Cells(1, 5) = "A二发ACE"
            .Cells(1, 7) = "A二发非ACE"
            .Cells(1, 9) = "B一发ACE"
            .Cells(1, 11) = "B一发非ACE"
            .Cells(1, 13) = "B二发ACE"
            .Cells(1, 15) = "B二发非ACE"
            .Cells(1, 17) = "A击球点"
            .Cells(1, 19) = "B击球点"
            
            Dim bb As New vecBall
            Call bb.init
            Call bb.combine(A1stServeAce)                               'A一发ACE
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 1) = ba.x
                  .Cells(i + 2, 2) = ba.y
            Next i
            Call bb.init
            Call bb.combine(A1stServeInWithoutAce)                      'A一发非ACE
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 3) = ba.x
                  .Cells(i + 2, 4) = ba.y
            Next i
            Call bb.init
            Call bb.combine(A2ndServeAce)                               'A二发ACE
            For i = 0 To bb.count - 1
                  Set ba = bb.pop_back()
                  .Cells(i + 2, 5) = ba.x
                  .Cells(i + 2, 6) = ba.y
            Next i
            Call bb.init
            Call bb.combine(A2ndServeInWithoutAce)                      'A二发非ACE
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
'      Debug.Print "中路In"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
'      Call ForShow.combine(B2ndServeLandingPointMediumFault)
'      Call ForShow.combine(B1stServeLandingPointMediumFault)
'      Debug.Print "中路out"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
''      Call ForShow.combine(B1stServeLandingPointMedium)
'      Call ForShow.combine(BHitLandingPointFault)
'      Debug.Print "B击球out落点"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
''      Call ForShow.init
'      Call ForShow.combine(B1stServeLandingPointMediumFault)
'      Call ForShow.combine(B2ndServeLandingPointMediumFault)
'      Debug.Print "A接发反手In落点"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
'      Call ForShow.combine(B1stServeLandingPointOuter)
'      Call ForShow.combine(B2ndServeLandingPointOuter)
'      Debug.Print "外角In"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
'      Call ForShow.combine(B1stServeLandingPointOuterFault)
'      Call ForShow.combine(B2ndServeLandingPointOuterFault)
'      Debug.Print "内角out"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next
'      Call ForShow.init
'      Call ForShow.combine(B1stServeLandingPointOtherFault)
'      Call ForShow.combine(B2ndServeLandingPointOtherFault)
'      Debug.Print "其他out"
'      For i = 1 To ForShow.count
'            ForShow.pop_back.show
'      Next


      ''''''''''''''''''''''''''''其他细分数据''''''''''''''''''''''''''''''''''
      With Sheets("raceOtherDetails")
            .Range(.Cells(4, 3), .Cells(32, 4)).Clear
            .Range(.Cells(4, 6), .Cells(32, 7)).Clear
            .Range(.Cells(4, 9), .Cells(32, 9)).Clear
            .Columns("C:D").HorizontalAlignment = xlCenter
            .Columns("F:G").HorizontalAlignment = xlCenter
            .Columns("I:J").HorizontalAlignment = xlCenter
            
            .Cells(4, 3) = A1stServePoint.count - A1stServeLet.count                            'A方一发总数
            .Cells(5, 3) = A1stServeAce.count                                                   'A方一发ACE
            .Cells(6, 3) = A1stServeLandingPointInner.count                                     'A方一发_内角In
            .Cells(7, 3) = A1stServeLandingPointMedium.count                                    'A方一发_中路In
            .Cells(8, 3) = A1stServeLandingPointOuter.count                                     'A方一发_外角In
            .Cells(9, 3) = A1stServeLandingPointOtherFault.count                                'A方一发其他Out
            .Cells(10, 3) = A1stServeLandingPointInnerFault.count                               'A方一发_内角Out
            .Cells(11, 3) = A1stServeLandingPointMediumFault.count                              'A方一发_中路Out
            .Cells(12, 3) = A1stServeLandingPointOuterFault.count                               'A方一发_外角Out
            .Cells(13, 3) = A1stServeWin.count                                                  'A方一发得分回合
            .Cells(14, 3) = A2ndServePoint.count - A2ndServeLet.count                           'A方二发总数
            .Cells(15, 3) = A2ndServeAce.count                                                  'A方二发ACE
            .Cells(16, 3) = A2ndServeLandingPointInner.count                                    'A方二发_内角In
            .Cells(17, 3) = A2ndServeLandingPointMedium.count                                   'A方二发_中路In
            .Cells(18, 3) = A2ndServeLandingPointOuter.count                                    'A方二发_外角In
            .Cells(19, 3) = A2ndServeLandingPointOtherFault.count                               'A方二发其他Out
            .Cells(20, 3) = A2ndServeLandingPointInnerFault.count                               'A方二发_内角Out
            .Cells(21, 3) = A2ndServeLandingPointMediumFault.count                              'A方二发_中路Out
            .Cells(22, 3) = A2ndServeLandingPointOuterFault.count                               'A方二发_外角Out
            .Cells(23, 3) = A2ndServeWin.count                                                  'A方二发得分回合

            .Cells(25, 3).FormulaR1C1 = "=sum(R4C3,R14C3)"                                      'A方发球总数
            .Cells(26, 3).FormulaR1C1 = "=sum(R6C3,R16C3)"                                      'A方发球内角in
            .Cells(27, 3).FormulaR1C1 = "=sum(R7C3,R17C3)"                                      'A方发球中路in
            .Cells(28, 3).FormulaR1C1 = "=sum(R8C3,R18C3)"                                      'A方发球外角in
            .Cells(29, 3).FormulaR1C1 = "=sum(R9C3:R12C3,R19C3:R22C3)"                          'A方发球out
            .Cells(30, 3).FormulaR1C1 = "=IfError(R26C3/(R26C3 +R20C3 +R10C3), 0)"              'A方内角成功率
            .Cells(30, 3).NumberFormatLocal = "0%"
            .Cells(31, 3).FormulaR1C1 = "=IfError(R27C3/(R27C3 +R21C3 +R11C3), 0)"              'A方中路成功率
            .Cells(31, 3).NumberFormatLocal = "0%"
            .Cells(32, 3).FormulaR1C1 = "=IfError(R28C3/(R28C3 +R22C3 +R12C3), 0)"              'A方外角成功率
            .Cells(32, 3).NumberFormatLocal = "0%"


            .Cells(4, 6) = AReturnPoint.count                                                   'A方接发总数
            .Cells(5, 6) = AReturnLandingPointEasy.count                                        'A方接发易
            .Cells(6, 6) = AReturnLandingPointNormal.count                                      'A方接发中
            .Cells(7, 6) = AReturnLandingPointHard.count                                        'A方接发难
            .Cells(8, 6) = AReturnBeingVolleyPoint.count                                        'A方接发被截击
            .Cells(9, 6).FormulaR1C1 = "=sum(R5C6,R8C6)"                                        'A方接发成功
            .Cells(10, 6) = AReturnForehandTotal.count                                          'A方接发正手
            .Cells(11, 6) = AReturnBackhandTotal.count                                          'A方接发反手
            .Cells(12, 6) = AReturnForehandIn.count                                             'A方接发正手in
            .Cells(13, 6) = AReturnBackhandIn.count                                             'A方接发反手in
            .Cells(14, 6) = AReturnLandingPointFault.count                                      'A方接发out
            .Cells(15, 6).FormulaR1C1 = "=IfError(R9C6/R4C6, 0)"                                'A方接发成功率
            .Cells(15, 6).NumberFormatLocal = "0%"

            .Cells(18, 6) = ABreakSucceed.count                                                 'A方破发得分
            .Cells(19, 6) = ABreakPoint.count                                                   'A方破发点
            .Cells(20, 6) = ANetNeerByWin.count                                                 'A方上网得分
            .Cells(21, 6) = ANetNeerByPoint.count                                               'A方上网次数
            .Cells(22, 6) = A1stServeLet.count                                                  'A一发Let
            .Cells(23, 6) = A2ndServeLet.count                                                  'A二发Let

            .Cells(25, 6).FormulaR1C1 = "=sum(R5C3，R15C3)"                                     'A方ACE
            .Cells(26, 6).FormulaR1C1 = "=sum(R19C3:R22C3)"                                     'A方双误
            .Cells(27, 6).FormulaR1C1 = "=IfError(sum(R6C3:R8C3)/R4C3, 0)"                      'A方一发成功率
            .Cells(27, 6).NumberFormatLocal = "0%"
            .Cells(28, 6).FormulaR1C1 = "=IfError(R13C3/sum(R6C3:R8C3), 0)"                     'A方一发胜率
            .Cells(28, 6).NumberFormatLocal = "0%"
            .Cells(29, 6).FormulaR1C1 = "=IfError(R23C3/sum(R16C3:R18C3), 0)"                   'A方二发胜率
            .Cells(29, 6).NumberFormatLocal = "0%"
            .Cells(30, 6).FormulaR1C1 = "=IfError(R18C6/R19C6, 0)"                              'A方破发得分率
            .Cells(30, 6).NumberFormatLocal = "0%"
            .Cells(31, 6).FormulaR1C1 = "=IfError(R20C6/R21C6, 0)"                              'A方网前得分率
            .Cells(31, 6).NumberFormatLocal = "0%"
            .Cells(32, 6) = AWinner.count                                                       'A方致胜分
            
            
            .Cells(4, 9) = AHitPoint.count                                                      'A方击球总数
            .Cells(5, 9) = AHitLandingPointEasy.count                                           'A方击球易
            .Cells(6, 9) = AHitLandingPointNormal.count                                         'A方击球中
            .Cells(7, 9) = AHitLandingPointHard.count                                           'A方击球难
            .Cells(8, 9) = AHitBeingVolleyPoint.count                                           'A方击球被截击
            .Cells(9, 9).FormulaR1C1 = "=sum(R5C9,R8C9)"                                        'A方击球成功
            .Cells(10, 9) = AHitForehandTotal.count                                             'A方击球正手
            .Cells(11, 9) = AHitBackhandTotal.count                                             'A方击球反手
            .Cells(12, 9) = AHitForehandIn.count                                                'A方击球正手In
            .Cells(13, 9) = AHitBackhandIn.count                                                'A方击球反手In
            .Cells(14, 9) = AHitLandingPointFault.count                                         'A方击球out
            .Cells(15, 9).FormulaR1C1 = "=IfError(R12C9/R10C9, 0)"                              'A击球正手成功率
            .Cells(15, 9).NumberFormatLocal = "0%"
            .Cells(16, 9).FormulaR1C1 = "=IfError(R13C9/R11C9, 0)"                              'A击球反手成功率
            .Cells(16, 9).NumberFormatLocal = "0%"
            .Cells(18, 9).FormulaR1C1 = "=sum(R4C9,R4C6)"                                       'A方击球总数(含接发)
            .Cells(19, 9).FormulaR1C1 = "=sum(R12C9:R12C6)"                                     'A方击球正手in(含接发)
            .Cells(20, 9).FormulaR1C1 = "=sum(R13C9:R13C6)"                                     'A方击球反手in(含接发)
            .Cells(21, 9).FormulaR1C1 = "=sum(R14C9:R14C6)"                                     'A方击球out(含接发)
            
            .Cells(25, 9) = AShortRoundWin.count                                                'A方短拍胜利回合
            .Cells(26, 9) = AMediumRoundWin.count                                               'A方中拍胜利回合
            .Cells(27, 9) = ALongRoundWin.count                                                 'A方长拍胜利回合
            .Cells(28, 9) = ABoutWin.count                                                      'A方总得分回合数
            
            .Cells(4, 4) = B1stServePoint.count - B1stServeLet.count                            'B方一发总数
            .Cells(5, 4) = B1stServeAce.count                                                   'B方一发ACE
            .Cells(6, 4) = B1stServeLandingPointInner.count                                     'B方一发_内角In
            .Cells(7, 4) = B1stServeLandingPointMedium.count                                    'B方一发_中路In
            .Cells(8, 4) = B1stServeLandingPointOuter.count                                     'B方一发_外角In
            .Cells(9, 4) = B1stServeLandingPointOtherFault.count                                'B方一发其他Out
            .Cells(10, 4) = B1stServeLandingPointInnerFault.count                               'B方一发_内角Out
            .Cells(11, 4) = B1stServeLandingPointMediumFault.count                              'B方一发_中路Out
            .Cells(12, 4) = B1stServeLandingPointOuterFault.count                               'B方一发_外角Out
            .Cells(13, 4) = B1stServeWin.count                                                  'B方一发得分回合
            .Cells(14, 4) = B2ndServePoint.count - B2ndServeLet.count                           'B方二发总数
            .Cells(15, 4) = B2ndServeAce.count                                                  'B方二发ACE
            .Cells(16, 4) = B2ndServeLandingPointInner.count                                    'B方二发_内角In
            .Cells(17, 4) = B2ndServeLandingPointMedium.count                                   'B方二发_中路In
            .Cells(18, 4) = B2ndServeLandingPointOuter.count                                    'B方二发_外角In
            .Cells(19, 4) = B2ndServeLandingPointOtherFault.count                               'B方二发其他Out
            .Cells(20, 4) = B2ndServeLandingPointInnerFault.count                               'B方二发_内角Out
            .Cells(21, 4) = B2ndServeLandingPointMediumFault.count                              'B方二发_中路Out
            .Cells(22, 4) = B2ndServeLandingPointOuterFault.count                               'B方二发_外角Out
            .Cells(23, 4) = B2ndServeWin.count                                                  'B方二发得分回合

            .Cells(25, 4).FormulaR1C1 = "=sum(R4C4,R14C4)"                                      'B方发球总数
            .Cells(26, 4).FormulaR1C1 = "=sum(R6C4,R16C4)"                                      'B方发球内角in
            .Cells(27, 4).FormulaR1C1 = "=sum(R7C4,R17C4)"                                      'B方发球中路in
            .Cells(28, 4).FormulaR1C1 = "=sum(R8C4,R18C4)"                                      'B方发球外角in
            .Cells(29, 4).FormulaR1C1 = "=sum(R9C4:R12C4,R19C4:R22C4)"                          'B方发球out
            .Cells(30, 4).FormulaR1C1 = "=IfError(R26C4/(R26C4 +R20C4 +R10C4), 0)"              'B方内角成功率
            .Cells(30, 4).NumberFormatLocal = "0%"
            .Cells(31, 4).FormulaR1C1 = "=IfError(R27C4/(R27C4 +R21C4 +R11C4), 0)"              'B方中路成功率
            .Cells(31, 4).NumberFormatLocal = "0%"
            .Cells(32, 4).FormulaR1C1 = "=IfError(R28C4/(R28C4 +R22C4 +R12C4), 0)"              'B方外角成功率
            .Cells(32, 4).NumberFormatLocal = "0%"


            .Cells(4, 7) = BReturnPoint.count                                                   'B方接发总数
            .Cells(5, 7) = BReturnLandingPointEasy.count                                        'B方接发易
            .Cells(6, 7) = BReturnLandingPointNormal.count                                      'B方接发中
            .Cells(7, 7) = BReturnLandingPointHard.count                                        'B方接发难
            .Cells(8, 7) = BReturnBeingVolleyPoint.count                                        'B方接发被截击
            .Cells(9, 7).FormulaR1C1 = "=sum(R5C7,R8C7)"                                        'B方接发成功
            .Cells(10, 7) = BReturnForehandTotal.count                                          'B方接发正手
            .Cells(11, 7) = BReturnBackhandTotal.count                                          'B方接发反手
            .Cells(12, 7) = BReturnForehandIn.count                                             'B方接发正手in
            .Cells(13, 7) = BReturnBackhandIn.count                                             'B方接发反手in
            .Cells(14, 7) = BReturnLandingPointFault.count                                      'B方接发out
            .Cells(15, 7).FormulaR1C1 = "=IfError(R9C7/R4C7, 0)"                                'B方接发成功率
            .Cells(15, 7).NumberFormatLocal = "0%"

            .Cells(18, 7) = BBreakSucceed.count                                                 'B方破发得分
            .Cells(19, 7) = BBreakPoint.count                                                   'B方破发点
            .Cells(20, 7) = BNetNeerByWin.count                                                'B方上网得分
            .Cells(21, 7) = BNetNeerByPoint.count                                                  'B方上网次数
            .Cells(22, 7) = B1stServeLet.count                                                  'B一发Let
            .Cells(23, 7) = B2ndServeLet.count                                                  'B二发Let

            .Cells(25, 7).FormulaR1C1 = "=sum(R5C4,R15C4)"                                      'B方ACE
            .Cells(26, 7).FormulaR1C1 = "=sum(R19C4:R22C4)"                                     'B方双误
            .Cells(27, 7).FormulaR1C1 = "=IfError(sum(R6C4:R8C4)/R4C4, 0)"                      'B方一发成功率
            .Cells(27, 7).NumberFormatLocal = "0%"
            .Cells(28, 7).FormulaR1C1 = "=IfError(R13C4/sum(R6C4:R8C4), 0)"                     'B方一发胜率
            .Cells(28, 7).NumberFormatLocal = "0%"
            .Cells(29, 7).FormulaR1C1 = "=IfError(R23C4/sum(R16C4:R18C4), 0)"                   'B方二发胜率
            .Cells(29, 7).NumberFormatLocal = "0%"
            .Cells(30, 7).FormulaR1C1 = "=IfError(R18C7/R19C7, 0)"                              'B方破发得分率
            .Cells(30, 7).NumberFormatLocal = "0%"
            .Cells(31, 7).FormulaR1C1 = "=IfError(R20C7/R21C7, 0)"                              'B方网前得分率
            .Cells(31, 7).NumberFormatLocal = "0%"
            .Cells(32, 7) = BWinner.count                                                       'B方致胜分
            
            
            .Cells(4, 10) = BHitPoint.count                                                      'B方击球总数
            .Cells(5, 10) = BHitLandingPointEasy.count                                           'B方击球易
            .Cells(6, 10) = BHitLandingPointNormal.count                                         'B方击球中
            .Cells(7, 10) = BHitLandingPointHard.count                                           'B方击球难
            .Cells(8, 10) = BHitBeingVolleyPoint.count                                           'B方击球被截击
            .Cells(9, 10).FormulaR1C1 = "=sum(R5C10,R8C10)"                                      'B方击球成功
            .Cells(10, 10) = BHitForehandTotal.count                                             'B方击球正手
            .Cells(11, 10) = BHitBackhandTotal.count                                             'B方击球反手
            .Cells(12, 10) = BHitForehandIn.count                                                'B方击球正手In
            .Cells(13, 10) = BHitBackhandIn.count                                                'B方击球反手In
            .Cells(14, 10) = BHitLandingPointFault.count                                         'B方击球out
            .Cells(15, 10).FormulaR1C1 = "=IfError(R12C10/R10C10, 0)"                            'B击球正手成功率
            .Cells(15, 10).NumberFormatLocal = "0%"
            .Cells(16, 10).FormulaR1C1 = "=IfError(R13C10/R11C10, 0)"                            'B击球反手成功率
            .Cells(16, 10).NumberFormatLocal = "0%"
            .Cells(18, 10).FormulaR1C1 = "=sum(R4C10,R4C7)"                                      'A方击球总数(含接发)
            .Cells(19, 10).FormulaR1C1 = "=sum(R12C10,R12C7)"                                    'A方击球正手in(含接发)
            .Cells(20, 10).FormulaR1C1 = "=sum(R13C10,R13C7)"                                    'A方击球反手in(含接发)
            .Cells(21, 10).FormulaR1C1 = "=sum(R14C10,R14C7)"                                    'A方击球out(含接发)
            
            .Cells(25, 10) = BShortRoundWin.count                                                'B方短拍胜利回合
            .Cells(26, 10) = BMediumRoundWin.count                                               'B方中拍胜利回合
            .Cells(27, 10) = BLongRoundWin.count                                                 'B方长拍胜利回合
            .Cells(28, 10) = BBoutWin.count                                                      'B方总得分回合数
            
      End With
      
End Sub

Sub breakPointGrabber(j%, gameMode%, isWinnerServe As Boolean, ba As ball, _
                  WinnerBout%, WinnerGame%, WinnerBreakPoint As vecBall, WinnerBreakSucceed As vecBall, _
                  serveFlag%, Winner1stServeWin As vecBall, Winner2ndServeWin As vecBall, WinnerBoutWin As vecBall, _
                  LoserBout%, LoserGame%)
      If gameMode = 4 Or gameMode = 6 Then                                                      '非抢七模式
            If WinnerGame = gameMode And WinnerGame = LoserGame Then                            '正在抢七局
                  If WinnerBout >= 7 And WinnerBout >= LoserBout + 2 Then
                        LoserBout = 0
                        LoserGame = 0
                        WinnerBout = 0
                        WinnerGame = 0
                  End If
            Else                                                                                '正在普通局
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




