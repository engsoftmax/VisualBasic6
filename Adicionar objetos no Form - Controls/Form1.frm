VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Adicionar Controls"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Licenses RichTextCtrl"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Adicionar um CommandButton com eventos"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remover objetos"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adicionar objetos"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents F_CommandButton As CommandButton
Attribute F_CommandButton.VB_VarHelpID = -1

Private Sub Command1_Click()
  Dim P_T1 As TextBox
  Dim P_T2 As TextBox

  Set P_T1 = Controls.Add("VB.TextBox", "Texto1", Form1)
  Set P_T2 = Controls.Add("VB.TextBox", "Texto2", Form1)
  
  P_T1.Move 24, 8, 184, 20
  P_T1.Text = "Caixa de Texto 1"
  P_T1.Visible = True
  
  P_T2.Move 24, 32, 184, 20
  P_T2.Text = "Caixa de Texto 2"
  P_T2.Visible = True
    
  Set P_T1 = Nothing
  Set P_T2 = Nothing
  
  'ou utilize a forma abaixo
  Form1.Controls.Add "VB.TextBox", "Texto3"
  
  Form1!Texto3.Move 24, 56, 184, 20
  Form1!Texto3.Text = "Caixa de Texto 3"
  Form1!Texto3.Visible = True
End Sub

Private Sub Command2_Click()
  Controls.Remove Form1.Controls("Texto1")
  Controls.Remove Form1.Controls("Texto2")
  Controls.Remove Form1.Controls("Texto3")
  
  'ou
  
  'Controls.Remove "Texto1"
  'Controls.Remove "Texto2"
  'Controls.Remove "Texto3"
End Sub

Private Sub Command3_Click()
  Set F_CommandButton = Form1.Controls.Add("VB.CommandButton", "NomeDoBotao")

  F_CommandButton.Caption = "&Clique aqui"
  F_CommandButton.Move Command3.Left, Command3.Top + 32, Command3.Width, Command3.Height
  F_CommandButton.Visible = True
End Sub

Private Sub Command4_Click()
  'Dim P_L As LicenseInfo
  
  'For Each P_L In Licenses
  '  Print P_L.ProgId & " - " & P_L.LicenseKey
  'Next P_L
  
  Licenses.Add "RichText.RichTextCtrl", " qhj ZtuQha;jdfn[iaetr "
  Form1.Controls.Add "RichText.RichTextCtrl", "RichText1"
  Licenses.Remove "RichText.RichTextCtrl"
  
  Form1.Controls("RichText1").Move Command4.Left, 8, Command4.Width, Command4.Top - 16
  Form1.Controls("RichText1").Visible = True
End Sub

Private Sub F_CommandButton_Click()
  MsgBox "Este botão foi criado usando:" & vbCr & vbCr & "Dim WithEvents F_CommandButton As CommandButton"
End Sub

