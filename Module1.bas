Attribute VB_Name = "Module1"
Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public DS As New DS_Engine
Public Type Player
PlaneTYPE As Integer
score As Long
End Type: Public Pl As Player, Comp As Player
Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Public R1 As RECT, R2 As RECT, R3 As RECT, Level As Integer

Function DetectCollision() As Boolean
R1.Top = p(1).Y - 16
R1.Bottom = p(1).Y + 16
R1.Left = p(1).X - 16
R1.Right = p(1).X + 16

R2.Top = p(2).Y - 16
R2.Bottom = p(2).Y + 16
R2.Left = p(2).X - 16
R2.Right = p(2).X + 16

IntersectRect R3, R1, R2
If R3.Right > 5 Then
DetectCollision = True
DS.PlaySound2 3
Else
If R3.Bottom > 5 Then
DetectCollision = True
DS.PlaySound2 3
Else
DetectCollision = False
End If
End If
R3.Right = 0
R3.Left = 0
R3.Top = 0
R3.Bottom = 0
End Function
Public Function RndRange(ByVal intMin As Integer, ByVal intMax As Integer)
RndRange = Int(Rnd * (intMax - intMin + 1)) + intMin
End Function
