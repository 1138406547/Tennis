VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vecBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_num As Integer
Private m_ball() As New ball

Private Sub Class_Initialize()
      m_num = 0
      ReDim m_ball(50)
End Sub
Public Sub init()
      m_num = 0
      ReDim m_ball(50)
End Sub

Property Get count() As Integer
      count = m_num
End Property

Public Sub push_back(ba As ball)                 'like c++ vector's push_back
      If m_num > UBound(m_ball) Then
            ReDim Preserve m_ball(m_num + 50)
      End If
      Call m_ball(m_num).clone(ba)
      m_num = m_num + 1
End Sub

Public Function pop_back() As ball              'like c++ vector's pop_back
      Dim tmp As New ball
      If m_num > 0 Then
            Call tmp.clone(m_ball(m_num - 1))
            m_num = m_num - 1
      End If
      Set pop_back = tmp
End Function

Public Function getBall(i As Integer) As ball               'make private ball array public
      If i >= LBound(m_ball) And i <= UBound(m_ball) Then
            Set getBall = m_ball(i)
      Else
            Set getBall = Nothing
      End If
End Function

Public Sub combine(vcb As vecBall)
      Dim i%
      If vcb.count > 0 Then
            ReDim Preserve m_ball(m_num + vcb.count - 1)
            For i = 0 To vcb.count - 1
                  Call m_ball(m_num + i).clone(vcb.getBall(i))
            Next
            m_num = m_num + vcb.count
      End If
End Sub


'''''functions below is just for test

Public Sub showAll()                            'show all balls in this vector
      Dim i
      For i = 0 To m_num - 1
            m_ball(i).show
      Next
End Sub

Public Sub showTop()                            'show the toppist ball
      m_ball(m_num - 1).show
End Sub

