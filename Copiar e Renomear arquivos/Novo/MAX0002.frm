VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Copiar e Renomear"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   2190
      Width           =   975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Iniciar"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Prefixo:"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1965
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim Tmp1 As Long
  Dim Tmp2 As String
  Dim Tmp3 As String
  Dim Tmp4 As String
  Dim Tmp5 As String
  
  Tmp2 = Dir1.Path & "\*.*"
  Tmp4 = Dir1.Path & "\"
  Tmp5 = Dir1.Path & "\Novo"
  
  Tmp1 = 0
  
  If Dir(Tmp5, vbDirectory) <> "" Then
    MsgBox "O diretório '" & UCase(Tmp5) & "' já existe!", vbInformation, "Impossível Continuar"
    Exit Sub
  Else
    MkDir Tmp5
  End If
  
  Tmp5 = Tmp5 & "\" & Text1.Text
  
  Screen.MousePointer = 11
  Refresh
  
  Tmp2 = Dir(Tmp2)
  Do Until Tmp2 = ""
    Tmp1 = Tmp1 + 1
    
    Tmp3 = Tmp5 & Format(Tmp1, "0000") & LCase(Mid(Tmp2, InStrRev(Tmp2, ".")))
    
    Tmp2 = Tmp4 & Tmp2
    FileCopy Tmp2, Tmp3
    
    Command1.Caption = Tmp1
    Command1.Refresh
    
    Tmp2 = Dir
  Loop
  
  Screen.MousePointer = 0
  
  Unload Me
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()

End Sub


Private Sub Form_Unload(Cancel As Integer)
  End
End Sub


