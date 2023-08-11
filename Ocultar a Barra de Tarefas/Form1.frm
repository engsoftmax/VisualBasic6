VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Barra de tarefas"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Exibir a barra de tarefas"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ocultar a barra de tarefas"
      Height          =   495
      Left            =   120
      TabIndex        =   0
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

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40

Private Sub Command1_Click()
  Dim P_hWnd As Long
  
  P_hWnd = FindWindow("Shell_traywnd", "")
  SetWindowPos P_hWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW
End Sub


Private Sub Command2_Click()
  Dim P_hWnd As Long
  
  P_hWnd = FindWindow("Shell_traywnd", "")
  SetWindowPos P_hWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW
End Sub


