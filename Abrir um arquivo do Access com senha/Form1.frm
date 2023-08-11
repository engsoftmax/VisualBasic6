VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Access com senha de abertura"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "Form1.frx":0000
      Left            =   600
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Abrir Banco de Dados com senha"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim P_DB As Database
  Dim P_RS As Recordset
  Dim P_FLD As Field
  
  Set P_DB = Workspaces(0).OpenDatabase(App.Path & "\ComSenha.mdb", False, False, "MS Access;PWD=1234")
  Set P_RS = P_DB.OpenRecordset("Cadastro", dbOpenTable)
  Set P_FLD = P_RS.Fields("Produto")
  
  P_RS.Index = "Codigo"
  
  If P_RS.RecordCount > 0 Then
    P_RS.MoveFirst
    
    List1.Clear
    
    Do Until P_RS.EOF
      List1.AddItem P_FLD.Value
      
      P_RS.MoveNext
    Loop
  End If
  
  P_RS.Close
  P_DB.Close
  
  Set P_FLD = Nothing
  Set P_RS = Nothing
  Set P_DB = Nothing
End Sub


