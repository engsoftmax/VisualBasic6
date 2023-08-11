VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Clique com o botão direito aqui - Opção Sobre"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim d As String

  d = SubClass(Form1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SetWindowLong Me.hWnd, GWL_WNDPROC, lProcOld
End Sub


