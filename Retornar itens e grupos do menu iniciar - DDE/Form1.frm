VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cItems 
      Caption         =   "&Items"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cGroups 
      Caption         =   "&Groups"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox tItems 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox tGroups 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.ComboBox ComItems 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.ComboBox ComGroups 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cGroups_Click()
  Dim sGroups As String
  Dim pos As Integer
  
  On Error GoTo GError
  
  tGroups.LinkMode = 0
  tGroups.LinkTopic = "Progman|Progman"
  tGroups.LinkMode = 2
  'tGroups.LinkItem = "Progman"
  tGroups.LinkItem = "groups"
  tGroups.LinkRequest
  
  'Parse groups that come back:
  sGroups = tGroups.Text
  pos = InStr(1, sGroups, Chr(13))
  While pos
    ComGroups.AddItem RTrim(Mid(sGroups, 1, pos - 1))
    sGroups = LTrim(Mid(sGroups, pos + 2))
    'The + 2 on the previous line gets past the line feed chr(10)
    pos = InStr(1, sGroups, Chr(13))
  Wend
  'Select first member in combo box:
  ComGroups.ListIndex = 1
  
GDone:
  tGroups.LinkMode = 0
  Exit Sub
  
GError:
  MsgBox "Error in getting groups"
  Resume GDone
End Sub


Private Sub cItems_Click()
  Dim sItems As String
  Dim pos As Integer

  On Error GoTo IError
  'Clear the combo box:
  ComItems.Clear
  If (Len(ComGroups.Text)) Then
    tItems.LinkMode = 0
    tItems.LinkTopic = "Progman|Progman"
    tItems.LinkMode = 2
    tItems.LinkItem = ComGroups.Text
    tItems.LinkRequest
    
    'Parse items that come back:
    sItems = tItems.Text
    pos = InStr(1, sItems, Chr(13))
    While pos
      ComItems.AddItem RTrim$(Mid$(sItems, 1, pos - 1))
      sItems = LTrim$(Mid$(sItems, pos + 2))
      'The + 2 on the previous line gets past the line feed chr(10)
      pos = InStr(1, sItems, Chr(13))
    Wend
  End If
  
  'Select first member in combo box:
  ComItems.ListIndex = 1
  
IDone:
  tItems.LinkMode = 0
  
  Exit Sub
IError:
  MsgBox "Error in getting items"
  Resume IDone
End Sub

