VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Vectors"
   ClientHeight    =   11805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11265
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11805
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2280
      Top             =   1920
   End
   Begin VB.PictureBox Display 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   600
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   593
      TabIndex        =   0
      Top             =   480
      Width           =   8895
   End
   Begin VB.PictureBox Buffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   8175
      Left            =   360
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   745
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   11175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectX, SelectY, Going, G, FPS

Private Sub Display_Click()
Cursor True
End
End Sub

Private Sub Form_Click()
Cursor True
End
End Sub

Private Sub Form_Load()
Form1.Left = 0
Form1.Top = 0
Form1.Width = Screen.Width
Form1.Height = Screen.Height
Display.Left = (Screen.Width - Display.Width) / 2
Display.Top = (Screen.Height - Display.Height) / 2
Form1.ForeColor = RGB(255, 255, 255)
Cursor False
R1 = 255
G1 = 255
B1 = 0
R2 = 255
G2 = 0
B2 = 0
Display.AutoRedraw = True
SetWave
DrawFrame
DoEvents
Display.AutoRedraw = False
Form1.Visible = True
Do
DoEvents
RotateVecs Sin(G / 4) * 8, Sin(G) + 1, 1
TotZ = 60
DrawFrameB 'DrawPic '
G = G + 0.5
R1 = (Sin(G / 6) * 128 + 128)
G1 = 0
B1 = (Cos(G / 4) * 128 + 128)
R2 = (Sin(G / 2) * 128 + 128)
G2 = (Cos(G / 2) * 128 + 128)
B2 = 0
For i = -9 To 10
For X = -9 To 10
    vector(i + 10, X + 10).X = (Sin(G + Sqr((vector(i + 10, X + 10).X ^ 2) + (vector(i + 10, X + 10).Y ^ 2) + (vector(i + 10, X + 10).Z ^ 2)) / 20) * 30) 'Sin((G / 20)) + G) * 50
    vector(i + 10, X + 10).Y = i * 15
    vector(i + 10, X + 10).Z = X * 15
Next X
Next i
FPS = FPS + 1
Loop
Cursor True
End Sub

Private Sub Timer1_Timer()
Cls
Print "Grid Morph - FPS: " & FPS & " - Time: " & Time
FPS = 0
End Sub
