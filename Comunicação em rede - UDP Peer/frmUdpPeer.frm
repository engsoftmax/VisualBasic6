VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUdpPeer 
   Caption         =   "UDP Peer"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRemotePort 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Text            =   "localhost"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin MSWinsockLib.Winsock udpPeer 
      Left            =   1080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "RemotePort:"
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "RemoteHost:"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   930
   End
   Begin VB.Label lblRecebido 
      AutoSize        =   -1  'True
      Caption         =   "Recebido:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enviar:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblLocalPort 
      AutoSize        =   -1  'True
      Caption         =   "LocalPort: 0"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmUdpPeer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
  'Pode-se usar LocalPort fixo ou deixar o sistema
  'pegar o proximo.
  'udpPeer.Bind 1001
  
  udpPeer.Bind
  lblLocalPort.Caption = "LocalPort: " & udpPeer.LocalPort
End Sub


Private Sub Form_Unload(Cancel As Integer)
  If udpPeer.State <> sckClosed Then
    udpPeer.Close
  End If
End Sub

Private Sub txtRemoteHost_GotFocus()
  AutoSelect txtRemoteHost
End Sub


Private Sub txtRemoteHost_LostFocus()
  udpPeer.RemoteHost = txtRemoteHost.Text
End Sub


Private Sub txtRemotePort_GotFocus()
  AutoSelect txtRemotePort
End Sub


Private Sub txtRemotePort_LostFocus()
  udpPeer.RemotePort = Val(txtRemotePort.Text)
End Sub


Private Sub txtSend_Change()
  On Error Resume Next
  
  udpPeer.SendData txtSend.Text
  
  If Err.Number <> 0 Then
    MsgBox Err.Description, vbInformation, Caption
  End If
  
  On Error GoTo 0
End Sub



Public Sub AutoSelect(P_TextBox As TextBox)
  P_TextBox.SelStart = 0
  P_TextBox.SelLength = Len(P_TextBox.Text)
End Sub

Private Sub txtSend_GotFocus()
  AutoSelect txtSend
End Sub


Private Sub udpPeer_DataArrival(ByVal bytesTotal As Long)
  Dim P_Data As String
  
  udpPeer.GetData P_Data
  txtOutput.Text = P_Data
End Sub

