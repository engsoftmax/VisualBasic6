VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Somente UCase ou LCase"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Lowercase Only"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Uppercase Only"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub ChangeCase(P_TextBox As Control, P_UpperCase As Boolean)
  Const GWL_STYLE = (-16)
  Const ES_UPPERCASE = &H8&
  Const ES_LOWERCASE = &H10&

  Dim EditStyle As Long
  
  EditStyle = GetWindowLong(P_TextBox.hwnd, GWL_STYLE)
  
  If P_UpperCase = True Then
    If (EditStyle And ES_LOWERCASE) Then
      EditStyle = EditStyle Xor ES_LOWERCASE
    End If
    
    EditStyle = EditStyle Or ES_UPPERCASE
  End If
  
  If P_UpperCase = False Then
    If (EditStyle And ES_UPPERCASE) Then
      EditStyle = EditStyle Xor ES_UPPERCASE
    End If
    
    EditStyle = EditStyle Or ES_LOWERCASE
  End If
  
  SetWindowLong P_TextBox.hwnd, GWL_STYLE, EditStyle
End Sub
Private Sub Command1_Click()
  Text1.Text = ""
  ChangeCase Text1, True
  Text1.SetFocus
End Sub


Private Sub Command2_Click()
  Text1.Text = ""
  ChangeCase Text1, False
  Text1.SetFocus
End Sub


