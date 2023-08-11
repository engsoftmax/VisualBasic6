VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ler e gravar no registro"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Gravar Valor"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ler Valor"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Este exemplo acelera a exibição dos submenus"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Command1.Caption = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\desktop", "MenuShowDelay")
End Sub


Private Sub Command2_Click()
  UpdateKey HKEY_CURRENT_USER, "Control Panel\desktop", "MenuShowDelay", "16"
End Sub


