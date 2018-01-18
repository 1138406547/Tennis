VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "multiZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_num As Integer                  'points to the next empty element
Private m_zone() As New zone

Private Sub Class_Initialize()
      m_num = 0
      ReDim m_zone(3)
End Sub

Public Sub push_back(z As zone)                 'like c++ vector's push_back
      If m_num > UBound(m_zone) Then
            ReDim Preserve m_zone(m_num * 2)
      End If
      Call m_zone(m_num).clone(z)
      m_num = m_num + 1
End Sub

Public Sub Clear()
      m_num = 0
      ReDim m_zone(3)
End Sub

Public Function count() As Integer
      count = m_num
End Function

Public Function isContain(ba As ball) As Boolean
      Dim i%
      For i = 0 To m_num - 1
            If m_zone(i).isContain(ba) Then
                  isContain = True
                  Exit Function
            End If
      Next
            isContain = False
End Function

Public Sub mirrorX()
      Dim z%, s%, y%, x%, i%
      For i = 0 To m_num - 1
            z = m_zone(i).left
            s = m_zone(i).down * (-1)
            y = m_zone(i).right
            x = m_zone(i).up * (-1)
            Call m_zone(i).init(z, s, y, x)
      Next
End Sub
Public Sub mirrorY()
      Dim z%, s%, y%, x%, i%
      For i = 0 To m_num - 1
            z = m_zone(i).right * (-1)
            s = m_zone(i).up
            y = m_zone(i).left * (-1)
            x = m_zone(i).down
            Call m_zone(i).init(z, s, y, x)
      Next
End Sub
Public Sub mirrorZ()
      Dim z%, s%, y%, x%, i%
      For i = 0 To m_num - 1
            z = m_zone(i).right * (-1)
            s = m_zone(i).down * (-1)
            y = m_zone(i).left * (-1)
            x = m_zone(i).up * (-1)
            Call m_zone(i).init(z, s, y, x)
      Next
End Sub

Public Function getZone(i As Integer) As zone
      If i >= LBound(m_zone) And i <= UBound(m_zone) Then
            Set getZone = m_zone(i)
      Else
            Set getZone = Nothing
      End If
End Function

Public Sub clone(mtz As multiZone)
      Dim i%
      ReDim m_zone(mtz.count)
      For i = 0 To mtz.count - 1
            Call m_zone(i).clone(mtz.getZone(i))
      Next
      m_num = mtz.count
End Sub

Public Sub combine(mtz As multiZone)
      Dim i%
      If mtz.count > 0 Then
            ReDim Preserve m_zone(m_num + mtz.count - 1)
            For i = 0 To mtz.count - 1
                  Call m_zone(m_num + i).clone(mtz.getZone(i))
            Next
            m_num = m_num + mtz.count
      End If
End Sub

''''' functions below is just for test


'Private Sub Class_Terminate()
'      Debug.Print "multiZone.terminate"
'End Sub

