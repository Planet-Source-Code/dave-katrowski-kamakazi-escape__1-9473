Attribute VB_Name = "modTileBG"
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020

Sub TileBack(Dest As Object, PB As PictureBox, Xoffset As Long, Yoffset As Long)
Dim i As Integer, ii As Integer, sm As Integer

sm = Dest.ScaleMode 'Save Current ScaleMode

Dest.ScaleMode = 3: PB.ScaleMode = 3

'Draw Loop:
For i = 0 To Dest.ScaleWidth * 5 Step PB.ScaleWidth
    For ii = 0 To Dest.ScaleHeight * 5 Step PB.ScaleHeight
        'Paint Tile
        BitBlt Dest.hDC, i - (2 * Xoffset), ii - (2 * Yoffset), PB.ScaleWidth, PB.ScaleHeight, PB.hDC, 0, 0, SRCCOPY
    Next
Next

Dest.ScaleMode = sm 'Set To Old ScaleMode
End Sub
