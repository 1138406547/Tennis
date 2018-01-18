VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_left As Integer
Private m_up As Integer
Private m_right As Integer
Private m_down As Integer

Property Get left()
      left = m_left
End Property
Property Get up()
      up = m_up
End Property
Property Get right()
      right = m_right
End Property
Property Get down()
      down = m_down
End Property

Public Sub init(z%, s%, y%, x%)
      m_left = z
      m_up = s
      m_right = y
      m_down = x
End Sub

Public Sub clone(z As zone)
      With z
            m_left = .left
            m_up = .up
            m_right = .right
            m_down = .down
      End With
End Sub

Public Function isContain(ba As ball) As Boolean
      With ba
            isContain = (.x >= m_left And .x <= m_right And .y <= m_up And .y >= m_down)
      End With
End Function

Public Sub mirrorX()
      Dim tmp%
      tmp = m_up
      m_up = m_down * (-1)
      m_down = tmp * (-1)
End Sub

Public Sub mirrorY()
      Dim tmp%
      tmp = m_right
      m_right = m_left * (-1)
      m_left = tmp * (-1)
End Sub

Public Sub mirrorZ()
      Dim tmp%
      tmp = m_up
      m_up = m_down * (-1)
      m_down = tmp * (-1)
      tmp = m_right
      m_right = m_left * (-1)
      m_left = tmp * (-1)
End Sub


'''''functions below is just for test
Public Sub show()
      Debug.Print m_left & "," & m_up & "," & m_right & "," & m_down
End Sub

'Private Sub Class_Initialize()
'      Debug.Print "zone.initialize"
'End Sub
'
'Private Sub Class_Terminate()
'      Debug.Print "zone.terminate"
'End Sub

