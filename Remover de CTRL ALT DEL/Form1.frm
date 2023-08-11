VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "App.Title"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "&Hiding Your Program in the Ctrl-Alt-Del list"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Private Const RSP_SIMPLE_SERVICE = 1
Private Const RSP_UNREGISTER_SERVICE = 0

Public Sub MakeMeService()
  Dim Pid As Long
  Dim Regserv As Long
  
  Pid = GetCurrentProcessId()
  Regserv = RegisterServiceProcess(Pid, RSP_SIMPLE_SERVICE)
End Sub

Public Sub UnMakeMeService()
  Dim Pid As Long
  Dim Regserv As Long
  
  Pid = GetCurrentProcessId()
  Regserv = RegisterServiceProcess(Pid, RSP_UNREGISTER_SERVICE)
End Sub

Private Sub Check1_Click()
  If Check1.Value = vbChecked Then
    MakeMeService
  Else
    UnMakeMeService
  End If
End Sub

Private Sub Form_Load()
  Caption = App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Check1.Value = vbChecked Then
    Check1.Value = vbUnchecked
  End If
End Sub


