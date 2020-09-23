Attribute VB_Name = "vect"
Declare Function Polygon& Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long)
Declare Function SetPolyFillMode& Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long)
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Type V
    X As Long
    Y As Long
    Z As Long
End Type
Type POINTAPI
        X As Long
        Y As Long
End Type

Global vector(20, 20) As V, Rotated As V, TotX, TotY, TotZ, VectorB(20, 20) As V
Global points(4) As POINTAPI, R1, R2, G1, G2, B1, B2, RI, GI, BI
Global Const SRCCOPY = &HCC0020

Sub FillPoly(Picture1 As PictureBox, P1X, P1Y, P2X, P2Y, P3X, P3Y, P4X, P4Y, Color)
Form1.Buffer.FillColor = Color
'Form1.Buffer.ForeColor = 0
points(0).X = P1X
points(0).Y = P1Y
points(1).X = P2X
points(1).Y = P2Y
points(2).X = P3X
points(2).Y = P3Y
points(3).X = P4X
points(3).Y = P4Y
Polygon Form1.Buffer.hDC, points(0), 4
End Sub

Sub FillPolyPic(X1, Y1, X2, Y2)
'BitBlt Form1.Buffer.hDC, X1, Y1, 15, 15, Form1.Picture1.hDC, X2, Y2, SRCCOPY
End Sub

Sub DrawPic()
hw = Form1.Buffer.ScaleWidth / 2
hh = Form1.Buffer.ScaleHeight / 2
Form1.Buffer.Cls
For i = 1 To 19
For X = 1 To 19
    clr = ((X * 10) / 190)
    FillPolyPic VectorB(X, i).X + hw, VectorB(X, i).Y + hh, i * 16, X * 16
Next X
Next i

BitBlt Form1.Display.hDC, 0, 0, Form1.Display.ScaleWidth, Form1.Display.ScaleHeight, Form1.Buffer.hDC, 0, 0, SRCCOPY
End Sub

Sub SetDefaults()
For i = -9 To 10
For X = -9 To 10
    vector(i + 10, X + 10).X = Int(Sin(Sin(i)) * 80)
    vector(i + 10, X + 10).Y = Int(Cos(Cos(i)) * 80)
    vector(i + 10, X + 10).Z = X * 15
Next X
Next i
TotX = 0
TotY = 0
TotZ = 0
End Sub

Sub SetRandZ()
For i = -9 To 10
For X = -9 To 10
    vector(i + 10, X + 10).X = X * 15
    vector(i + 10, X + 10).Y = i * 15
    vector(i + 10, X + 10).Z = Int(Rnd * 100)
Next X
Next i
TotX = 0
TotY = 0
TotZ = 0
End Sub

Sub SetWave()
For i = -9 To 10
For X = -9 To 10
    vector(i + 10, X + 10).X = Int(Sin(i) * 80)
    vector(i + 10, X + 10).Y = i * 15
    vector(i + 10, X + 10).Z = X * 15
Next X
Next i
TotX = 0
TotY = 0
TotZ = 0
End Sub

Sub SetCyllinder()
For i = -9 To 10
For X = -9 To 10
    vector(i + 10, X + 10).X = Int(Sin(i) * 80)
    vector(i + 10, X + 10).Y = Int(Cos(i) * 80)
    vector(i + 10, X + 10).Z = X * 15
Next X
Next i
TotX = 0
TotY = 0
TotZ = 0
End Sub

Sub DrawFrame()
hw = Form1.Buffer.ScaleWidth / 2
hh = Form1.Buffer.ScaleHeight / 2
Form1.Buffer.Cls
For i = 1 To 20
For X = 1 To 19
    Form1.Buffer.Line (vector(X, i).X + hw, vector(X, i).Y + hh)-(vector(X + 1, i).X + hw, vector(X + 1, i).Y + hh), clr
Next X
Next i

For X = 1 To 20
For i = 1 To 19
    Form1.Buffer.Line (vector(X, i).X + hw, vector(X, i).Y + hh)-(vector(X, i + 1).X + hw, vector(X, i + 1).Y + hh), clr
Next i
Next X

BitBlt Form1.Display.hDC, 0, 0, Form1.Display.ScaleWidth, Form1.Display.ScaleHeight, Form1.Buffer.hDC, 0, 0, SRCCOPY
End Sub

Sub DrawFrameB()
hw = Form1.Buffer.ScaleWidth / 2
hh = Form1.Buffer.ScaleHeight / 2
Form1.Buffer.Cls
For i = 1 To 19
For X = 1 To 19
    clr = ((X * 10) / 190)
    FillPoly Form1.Buffer, VectorB(X, i).X + hw, VectorB(X, i).Y + hh, VectorB(X + 1, i).X + hw, VectorB(X + 1, i).Y + hh, VectorB(X + 1, i + 1).X + hw, VectorB(X + 1, i + 1).Y + hh, VectorB(X, i + 1).X + hw, VectorB(X, i + 1).Y + hh, RGB(R1 + (R2 - R1) * clr, G1 + (G2 - G1) * clr, B1 + (B2 - B1) * clr)
Next X
Next i

BitBlt Form1.Display.hDC, 0, 0, Form1.Display.ScaleWidth, Form1.Display.ScaleHeight, Form1.Buffer.hDC, 0, 0, SRCCOPY
End Sub

Sub RotateVecs(XD, YD, ZD)
For i = 1 To 20
For X = 1 To 20
    VectorB(i, X).X = vector(i, X).X
    VectorB(i, X).Y = vector(i, X).Y
    VectorB(i, X).Z = vector(i, X).Z
Next X
Next i
TotX = TotX + XD
TotY = TotY + YD
TotZ = TotZ + ZD
XRadians = TotX * 3.141592653589 / 180
YRadians = TotY * 3.141592653589 / 180
ZRadians = TotZ * 3.141592653589 / 180
For i = 1 To 20
For X = 1 To 20
Rotated.X = VectorB(i, X).X
Rotated.Y = Cos(XRadians) * VectorB(i, X).Y - Sin(XRadians) * VectorB(i, X).Z
Rotated.Z = Sin(XRadians) * VectorB(i, X).Y + Cos(XRadians) * VectorB(i, X).Z

Rotated.X = Cos(YRadians) * Rotated.X - Sin(YRadians) * Rotated.Y
Rotated.Y = Sin(YRadians) * Rotated.X + Cos(YRadians) * Rotated.Y
Rotated.Z = Rotated.Z
    
Rotated.X = Cos(ZRadians) * Rotated.X - Sin(ZRadians) * Rotated.Z
Rotated.Y = Rotated.Y
Rotated.Z = Sin(ZRadians) * Rotated.X + Cos(ZRadians) * Rotated.Z
 
VectorB(i, X).X = Rotated.X
VectorB(i, X).Y = Rotated.Y
VectorB(i, X).Z = Rotated.Z
Next X
Next i
End Sub

Sub Cursor(Enabled As Boolean)
Dim Retcode
For i = 1 To 5000
Retcode = ShowCursor(Enabled)
Next i
End Sub
