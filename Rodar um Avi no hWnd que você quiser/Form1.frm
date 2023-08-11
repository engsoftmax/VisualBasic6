VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Reprodutor de vídeo AVI"
   ClientHeight    =   4710
   ClientLeft      =   2265
   ClientTop       =   1560
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4710
   ScaleWidth      =   4605
   Begin VB.CommandButton Command1 
      Caption         =   "Exibir vídeo"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2355
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   873
      _Version        =   393216
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      FileName        =   "d:\video\ap_01.avi"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Digite o caminho do arquivo avi que você deseja exibir"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  MMControl1.Command = "close"
  MMControl1.DeviceType = "avivideo"
  MMControl1.hWndDisplay = Picture1.hWnd
  MMControl1.FileName = Text1.Text
  MMControl1.Command = "open"
  MMControl1.Command = "prev"
  MMControl1.Command = "play"
End Sub

