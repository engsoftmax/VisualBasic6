VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Grid ==> MsFlexGrid"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Converter"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim Tmp1 As String
  Dim Tmp2 As Integer
  Dim Tmp3 As Integer
  Dim Tmp4 As String
  Dim Tmp5 As String
  Dim Tmp6 As String
  
  Tmp1 = Dir("C:\Tmp\Velho\*.FRM")
  Do Until Tmp1 = ""
    Tmp4 = "C:\Tmp\Velho\" & Tmp1
    Tmp5 = "C:\Tmp\Novo\" & Tmp1
    
    Tmp2 = FreeFile
    Open Tmp4 For Input As Tmp2
    Tmp3 = FreeFile
    Open Tmp5 For Output As Tmp3
    
    Do While Not EOF(Tmp2)
      Line Input #Tmp2, Tmp6
      If InStr(Tmp6, "Begin MSGrid.Grid") Then
        Tmp6 = "   Begin MSFlexGridLib.MSFlexGrid" & Right(Tmp6, 7)
        Print #Tmp3, Tmp6
        
        Do
          Line Input #Tmp2, Tmp6
          Select Case Trim(Left(Tmp6, 19))
          Case "Height", "Left", "TabIndex", "Top", "Width", "_ExtentX"
            Print #Tmp3, Tmp6
          Case "_ExtentY"
            Print #Tmp3, Tmp6
            Tmp6 = "      _Version        =   327680"
            Print #Tmp3, Tmp6
          Case "Rows", "Cols", "FixedCols", "FixedRows"
            Print #Tmp3, Tmp6
          Case "End"
            Exit Do
          End Select
        Loop
        Tmp6 = "      HighLight       =   0"
        Print #Tmp3, Tmp6
        Tmp6 = "      AllowUserResizing=   1"
        Print #Tmp3, Tmp6
        Tmp6 = "      Appearance      =   0"
        Print #Tmp3, Tmp6
        Tmp6 = "      FormatString    =   " & Chr(34) & Chr(34)
        Print #Tmp3, Tmp6
        Tmp6 = "   End"
        Print #Tmp3, Tmp6
      ElseIf InStr(Tmp6, "GRID32.OCX") Then  'Retirar Objeto
        Tmp6 = "Object = " & Chr(34) & "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0" & Chr(34) & "; " & Chr(34) & "MSFLXGRD.OCX" & Chr(34) & ""
        Print #Tmp3, Tmp6
      Else
        Print #Tmp3, Tmp6
      End If
    Loop
    
    Close Tmp2, Tmp3
    
    Tmp1 = Dir
  Loop
  
  Beep
  Unload Me
End Sub


