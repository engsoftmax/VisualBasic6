VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cores do sistema"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Atuais"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Sub Command1_Click()
  Dim Tmp1 As Integer
  
  For Tmp1 = 0 To 24
    Line (132, (Tmp1 * 15) + 15)-(142, (Tmp1 * 15) + 11 + 15), GetSysColor(Tmp1), BF
    Line (132, (Tmp1 * 15) + 15)-(142, (Tmp1 * 15) + 11 + 15), 0, B
  Next Tmp1
End Sub

