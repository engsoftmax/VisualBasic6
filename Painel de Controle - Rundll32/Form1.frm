VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Painel de Controle"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      Caption         =   "Configurações"
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Aparência"
      Height          =   195
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Proteção de Tela"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Segundo Plano"
      Height          =   195
      Left            =   2400
      TabIndex        =   0
      Top             =   2160
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   11
      Left            =   3240
      Picture         =   "Form1.frx":0000
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   2520
      Picture         =   "Form1.frx":030A
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   1800
      Picture         =   "Form1.frx":0614
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   1800
      Picture         =   "Form1.frx":091E
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line10 
      X1              =   2040
      X2              =   2280
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line9 
      X1              =   2040
      X2              =   2280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line8 
      X1              =   2040
      X2              =   2280
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line7 
      X1              =   2040
      X2              =   2280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line6 
      X1              =   2040
      X2              =   2040
      Y1              =   1920
      Y2              =   3000
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   1800
      Picture         =   "Form1.frx":0C28
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   3240
      Picture         =   "Form1.frx":0F32
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   2520
      Picture         =   "Form1.frx":123C
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   600
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   600
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   600
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   720
      Y2              =   3000
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   720
      Picture         =   "Form1.frx":1546
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   720
      Picture         =   "Form1.frx":1850
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   720
      Picture         =   "Form1.frx":1B5A
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   "Form1.frx":1E64
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":216E
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image1_Click(Index As Integer)
  Dim Tmp1 As String
  
  Select Case Index
  Case 0
    Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL access.cpl,,5"
  Case 1
    Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL access.cpl,,3"
  Case 2
    Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL access.cpl,,1"
  Case 3
    Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL access.cpl,,4"
  Case 4
    Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL access.cpl,,2"
  Case 5
    Tmp1 = "Rundll32.exe shell32.dll,Control_RunDLL main.cpl @2"
  Case 6
    Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl"
  Case 7
    Select Case True
    Case Option1.Value
      Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0"
    Case Option2.Value
      Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,1"
    Case Option3.Value
      Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2"
    Case Option4.Value
      Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3"
    End Select
  Case 8
    Tmp1 = "rundll32.exe shell32.dll,Control_RunDLL"
  Case 9
    Tmp1 = "Rundll32.exe shell32.dll,Control_RunDLL main.cpl @0"
  Case 10
    Tmp1 = "Rundll32.exe shell32.dll,Control_RunDLL main.cpl @1"
  Case 11
    Tmp1 = "Rundll32.exe shell32.dll,Control_RunDLL main.cpl @3"
  Case 12
    Tmp1 = "Rundll32.exe shell32.dll,Control_RunDLL main.cpl @4"
  End Select
  
  Shell Tmp1, vbNormalFocus
End Sub


