VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Abrir ComboBox"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Abrir"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const CB_SHOWDROPDOWN = &H14F

Private Sub Command1_Click()
  SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, 1, 0&
End Sub


Private Sub Command2_Click()
  SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, 0&, 0&
End Sub


Private Sub Form_Load()
  Combo1.AddItem "Banana"
  Combo1.AddItem "Abacate"
  Combo1.AddItem "Laranja"
  Combo1.AddItem "Melão"
  Combo1.AddItem "Uva"
  Combo1.AddItem "Goiaba"
End Sub


