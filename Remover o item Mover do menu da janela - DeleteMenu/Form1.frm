VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sem Mover"
   ClientHeight    =   885
   ClientLeft      =   2280
   ClientTop       =   2430
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   885
   ScaleWidth      =   4065
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Desabilita a opção MOVER quando se clica no ícone do programa."
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private Sub RemoveMenus(frm As Form, remove_restore As Boolean, remove_move As Boolean, remove_size As Boolean, remove_minimize As Boolean, remove_maximize As Boolean, remove_seperator As Boolean, remove_close As Boolean)
  Dim hMenu As Long
  
  ' Get the form's system menu handle.
  hMenu = GetSystemMenu(hwnd, False)
  
  If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
  If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
  If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
  If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
  If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
  If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
  If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

Private Sub Form_Load()
  RemoveMenus Me, False, True, False, False, False, False, False
  
  'A propriedade Moveable tem o mesmo efeito
End Sub

