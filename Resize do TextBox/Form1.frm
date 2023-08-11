VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resize TextBox"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_FRAMECHANGED = &H20   'The frame changed: send WM_NCCALCSIZE
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

Private Sub Command1_Click()
  Dim NewStyle As Long
  
  NewStyle = GetWindowLong(Text1.hwnd, GWL_STYLE)
  
  NewStyle = NewStyle Or WS_THICKFRAME
  
  NewStyle = SetWindowLong(Text1.hwnd, GWL_STYLE, NewStyle)
  
  SetWindowPos Text1.hwnd, hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub


