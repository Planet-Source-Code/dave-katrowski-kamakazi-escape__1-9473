Attribute VB_Name = "Particle"
Public Type Partical
X As Single
Y As Single
z As Single
a As Single
v As Long
End Type: Public p(1000) As Partical

Const Pi = 3.14159265358979 'Trig
Const PIdiv18 = Pi / 18
Public Sine(35) As Single, CoSn(35) As Single

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCAND = &H8800C6
Const SRCPAINT = &HEE0086
Const SRCCOPY = &HCC0020

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Keys(0 To 4) As Integer 'Add more as needed

Sub Math_BTT()
For i = 0 To 35
Sine(i) = Sin(i * PIdiv18)
CoSn(i) = Cos(i * PIdiv18)
Next
End Sub

Sub Move(ID As Integer, Direction As String, Vel As Long, TurningFriction As Single, Road_Air_Friction As Single)
Select Case Direction
Case "Left"
If p(ID).v > 0 Then
p(ID).a = p(ID).a + (Rnd * 1): If p(ID).a > 35 Then p(ID).a = 0
ElseIf p(ID).v < 0 Then
p(ID).a = p(ID).a - (Rnd * 1): If p(ID).a < 0 Then p(ID).a = 35
End If
If p(ID).v > 0 Then
p(ID).v = p(ID).v - TurningFriction
ElseIf p(ID).v < 0 Then
p(ID).v = p(ID).v + TurningFriction
End If
Case "Right"
If p(ID).v > 0 Then
p(ID).a = p(ID).a - (Rnd * 1): If p(ID).a < 0 Then p(ID).a = 35
ElseIf p(ID).v < 0 Then
p(ID).a = p(ID).a + (Rnd * 1): If p(ID).a > 35 Then p(ID).a = 0
End If
If p(ID).v > 0 Then
p(ID).v = p(ID).v - TurningFriction
ElseIf p(ID).v < 0 Then
p(ID).v = p(ID).v + TurningFriction
End If
Case "Forward"
p(ID).v = p(ID).v + Vel
Case "Reverse"
p(ID).v = p(ID).v - Vel
Case Else: 'Coast
End Select

p(ID).X = ((p(ID).X * 50) + (p(ID).v * Sine(p(ID).a))) / 50
p(ID).Y = ((p(ID).Y * 50) + (p(ID).v * CoSn(p(ID).a))) / 50

Select Case p(ID).v
Case Is > 0
If Not p(ID).v = 0 Then p(ID).v = p(ID).v - Road_Air_Friction
Case Is < 0
If Not p(ID).v = 0 Then p(ID).v = p(ID).v + Road_Air_Friction
End Select
End Sub

Sub DrawPIC(ID As Integer, SourceDC As Long, MaskDC As Long, DestDC As Long, W As Long, H As Long, Offset1 As Integer, offset2 As Integer)
If Not DetectCollision Then
BitBlt DestDC, p(ID).X + Offset1, p(ID).Y + offset2, W, H, MaskDC, 0, 0, SRCAND
Else
Pl.score = Pl.score - 5
End If
BitBlt DestDC, p(ID).X + Offset1, p(ID).Y + offset2, W, H, SourceDC, 0, 0, SRCPAINT
End Sub

Sub DrawPOINT(ID As Integer, Color As Long, PictureBox As Object)
PictureBox.PSet (p(ID).X, p(ID).Y), Color
End Sub



Sub SetXY(ID As Integer, X As Single, Y As Single)
p(ID).X = X: p(ID).Y = Y
End Sub
