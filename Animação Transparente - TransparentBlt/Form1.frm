VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2670
      Left            =   600
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   174
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   111
      TabIndex        =   1
      Top             =   1320
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   600
      Picture         =   "Form1.frx":47B4
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Somente 256 cores !!!"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim Tmp1 As Integer
  Dim Tmp2 As Single
  
  'Tmp2 = Timer
  
  'For Tmp1 = 1 To 10
    Call TransparentBlt(Picture2, Picture1.Picture, 10, 10, RGB(255, 0, 0))
  'Next Tmp1
  
  'Print Timer - Tmp2
  
  Picture2.Refresh
End Sub


Private Sub Command2_Click()
  SavePicture
End Sub


