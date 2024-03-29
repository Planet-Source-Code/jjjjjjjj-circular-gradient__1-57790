Attribute VB_Name = "Module1"
Option Explicit

Public Function CircularGradient(Pic As PictureBox, _
                                ByVal SCol As Long, _
                                ByVal ECol As Long)
Dim X As Integer
Dim XDist As Long
Dim YDist As Long
Dim DCol As Long
Dim DWidth As Long

Dim sR As Long, sG  As Long, sB  As Long
Dim eR   As Long, eG  As Long, eB  As Long
Dim Rincr As Double, Gincr As Double, Bincr As Double

XDist = Pic.Width / 2
YDist = Pic.Height / 2
DWidth = 1.5 * XDist / 100
GetRGB SCol, sR, sG, sB
GetRGB ECol, eR, eG, eB

Rincr = (eR - sR) / 100
Gincr = (eG - sG) / 100
Bincr = (eB - sB) / 100

Pic.DrawWidth = DWidth
Pic.AutoRedraw = True

On Error GoTo Handle
For X = 1 To 1.5 * XDist - 10
    DCol = RGB(sR, sG, sB)
    Pic.Circle (XDist, YDist), (1.5 * XDist - 1.2 * X * DWidth), DCol
    sR = sR + Rincr: sG = sG + Gincr: sB = sB + Bincr
Next X
Exit Function
Handle:
    If Err.Number = 5 Then
        Exit Function
    Else
        MsgBox Err.Description, vbCritical, Err.Number
    End If
End Function

'Gets the RGB values
Private Sub GetRGB(ByVal LngCol As Long, R As Long, G As Long, B As Long)
  R = LngCol Mod 256    'Red
  G = (LngCol And vbGreen) / 256 'Green
  B = (LngCol And vbBlue) / 65536 'Blue
End Sub
