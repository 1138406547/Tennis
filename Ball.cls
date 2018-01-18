VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_x As Integer
Private m_y As Integer
Private m_z As Integer
Private m_stat As Integer
Private m_speed As Integer
Private m_timeStamp As String

Property Let x(a%)
        m_x = a
End Property
Property Let y(a%)
        m_y = a
End Property
Property Let z(a%)
        m_z = a
End Property
Property Let stat(a%)
        m_stat = a
End Property
Property Let speed(a%)
        m_speed = a
End Property
Property Let timeStamp(a$)
        m_timeStamp = a
End Property

Property Get x() As Integer
        x = m_x
End Property
Property Get y() As Integer
        y = m_y
End Property
Property Get z() As Integer
        z = m_z
End Property
Property Get stat() As Integer
        stat = m_stat
End Property
Property Get speed() As Integer
        speed = m_speed
End Property
Property Get timeStamp() As String
        timeStamp = m_timeStamp
End Property

Public Sub clone(ba As ball)
      With ba
            m_x = .x
            m_y = .y
            m_z = .z
            m_stat = .stat
            m_speed = .speed
            m_timeStamp = .timeStamp
      End With
End Sub

Public Function isInside(z As multiZone) As Boolean
        isInside = z.isContain(Me)
End Function

Public Sub init(x%, y%, z%, stat%, Optional speed% = 0, Optional time$ = "")
      m_x = x
      m_y = y
      m_z = z
      m_stat = stat
      m_speed = speed
      m_timeStamp = time
End Sub

Public Sub show()
      Debug.Print m_x & "," & m_y & "," & m_z
End Sub

'Private Sub Class_Initialize()
'      Debug.Print "ball.initialize"
'End Sub
'
'Private Sub Class_Terminate()
'      Debug.Print "ball.terminate"
'End Sub
