VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Transparent Animation Demo"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   795
      Picture         =   "transparent.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   5
      Top             =   825
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5220
      Top             =   855
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   210
      Negotiate       =   -1  'True
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Tag             =   "0"
      Top             =   825
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   2745
      Picture         =   "transparent.frx":024A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1290
      Picture         =   "transparent.frx":0ACC
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      Top             =   225
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   210
      Negotiate       =   -1  'True
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Tag             =   "0"
      Top             =   225
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5220
      Top             =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   795
      Picture         =   "transparent.frx":6B0E
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1_Timer
End Sub

Private Sub Timer1_Timer()
With Picture1
.Item(0).Cls
.Item(0).PaintPicture .Item(1).Picture, 0, 0, , , , , , , vbSrcPaint
.Item(0).PaintPicture .Item(2).Picture, 0, 0, 32, 32, .Item(0).Tag * 32, 0, 32, 32, vbSrcAnd
.Item(0).Tag = .Item(0).Tag + 1
If .Item(0).Tag > 7 Then .Item(0).Tag = 0
End With
End Sub

Private Sub Timer2_Timer()
With Picture2
.Item(0).Cls
.Item(0).PaintPicture .Item(1).Picture, 0, 0, 32, 32, .Item(0).Tag * 32, 0, 32, 32, vbSrcPaint
.Item(0).PaintPicture .Item(2).Picture, 0, 0, 32, 32, .Item(0).Tag * 32, 0, 32, 32, vbSrcAnd
.Item(0).Tag = .Item(0).Tag + 1
If .Item(0).Tag > 3 Then .Item(0).Tag = 0
End With
End Sub
