VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Excluir a pasta C:\TestFolder"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar C:\*.* para C:\TestFolder"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim result As Long
  Dim fileop As SHFILEOPSTRUCT
  
  With fileop
    .hwnd = Me.hwnd
    .wFunc = FO_COPY
    '.pFrom = "C:\Teste1.Bmp" & vbNullChar & "C:\Teste2.TXT" & vbNullChar & vbNullChar
    .pFrom = "C:\*.*" & vbNullChar & vbNullChar
    'The directory or filename(s) to copy into terminated in 2 nulls.
    .pTo = "C:\TestFolder\" & vbNullChar & vbNullChar
    .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
  End With
  
  result = SHFileOperation(fileop)
  If result <> 0 Then
    'Operation failed
    MsgBox Err.LastDllError 'Msgbox the error that occurred in the API.
  Else
    If fileop.fAnyOperationsAborted <> 0 Then
      MsgBox "Operation Failed"
    End If
  End If
End Sub


Private Sub Command2_Click()
  Dim DelFileOp As SHFILEOPSTRUCT
  Dim result As Long
  
  With DelFileOp
    .hwnd = Me.hwnd
    .wFunc = FO_DELETE
    
    'Delete the files you just moved to C:\TestFolder.
    .pFrom = "C:\TestFolder" & vbNullChar & vbNullChar
    '.pFrom = "C:\TestFolder\*.*" & vbNullChar & vbNullChar
    '.pFrom = "C:\TestFolder\Teste1.TXT" & vbNullChar &  "C:\TestFolder\Teste2.TXT" & vbNullChar & vbNullChar
    
    'Allow undo--in other words, place the files into the Recycle Bin
    .fFlags = FOF_ALLOWUNDO
  End With
  
  result = SHFileOperation(DelFileOp)
  If result <> 0 Then
    'Operation failed
    MsgBox Err.LastDllError 'Msgbox the error that occurred in the API.
  Else
    If DelFileOp.fAnyOperationsAborted <> 0 Then
      MsgBox "Operation Failed"
    End If
  End If
End Sub


Private Sub Form_Load()

End Sub


