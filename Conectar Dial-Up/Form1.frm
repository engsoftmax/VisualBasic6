VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dial-Up"
   ClientHeight    =   1635
   ClientLeft      =   4275
   ClientTop       =   2070
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1635
   ScaleWidth      =   2955
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar-se"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Digite o nome da Conexão"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AbrirDialUp(Conexão As String)
  Dim X As Long
  
  X = Shell("RunDll32.exe rnaui.dll,RnaDial " & Conexão, 1)
  AppActivate X
End Sub

Private Sub Command1_Click()
  If Text1.Text = "" Then
    Text1.SetFocus
    MsgBox "Digite o nome de uma conexão Dial-Up", vbInformation, "Conexão"
  Else
    AbrirDialUp Text1.Text
  End If
End Sub


