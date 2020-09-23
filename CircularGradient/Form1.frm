VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Circular Gradient"
   ClientHeight    =   4020
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picE 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   8160
      ScaleHeight     =   1830
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   2040
      Width           =   405
   End
   Begin VB.PictureBox picS 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   8160
      ScaleHeight     =   1830
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   120
      Width           =   405
   End
   Begin VB.PictureBox picEnd 
      AutoSize        =   -1  'True
      Height          =   1860
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1800
      ScaleWidth      =   3345
      TabIndex        =   2
      Top             =   2040
      Width           =   3405
   End
   Begin VB.PictureBox picStart 
      AutoSize        =   -1  'True
      Height          =   1860
      Left            =   120
      Picture         =   "Form1.frx":13B42
      ScaleHeight     =   1800
      ScaleWidth      =   3345
      TabIndex        =   1
      Top             =   120
      Width           =   3405
   End
   Begin VB.PictureBox picGrad 
      AutoRedraw      =   -1  'True
      Height          =   3855
      Left            =   3720
      ScaleHeight     =   3795
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Startcol As Long
Dim EndCol As Long

Private Sub Form_Load()
    Startcol = 0: EndCol = RGB(255, 255, 255)
    CircularGradient picGrad, Startcol, EndCol
End Sub

Private Sub picEnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EndCol = picEnd.Point(X, Y)
    CircularGradient picGrad, Startcol, EndCol
    picS.BackColor = Startcol: picE.BackColor = EndCol
End Sub

Private Sub picStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Startcol = picStart.Point(X, Y)
    CircularGradient picGrad, Startcol, EndCol
    picS.BackColor = Startcol: picE.BackColor = EndCol
End Sub

