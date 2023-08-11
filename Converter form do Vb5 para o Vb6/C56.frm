VERSION 5.00
Begin VB.Form C56 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Converter Vb5 para Vb6"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Sai&r"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Converter"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Count = 0"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "C56"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  If Label1.Tag = 0 Then
    Beep
    Exit Sub
  End If
  
  Screen.MousePointer = 11
  Converter
  Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub


Private Sub Dir1_Change()
  Contar
End Sub

Private Sub Drive1_Change()
  Dim Tmp1 As String
  Tmp1 = Dir1.Path
  
  On Error GoTo RotError
  
  Dir1.Path = Left(Drive1.Drive, 1) & ":\"
  
  Exit Sub
RotError:
  Dir1.Path = Tmp1
  Drive1.Drive = Left(Tmp1, 1)
End Sub



Public Sub Contar()
  Dim Tmp1 As Integer
  Dim Tmp2 As String
  
  Tmp2 = Dir1.Path
  Select Case Right(Tmp2, 1)
  Case "\"
    Tmp2 = Tmp2 & "*.frm"
  Case Else
    Tmp2 = Tmp2 & "\*.frm"
  End Select
  
  Tmp1 = 0
  Tmp2 = Dir(Tmp2)
  Do Until Tmp2 = ""
    Tmp1 = Tmp1 + 1
    Tmp2 = Dir
  Loop
  Label1.Caption = "Count = " & Tmp1
  Label1.Tag = Tmp1
End Sub

Private Sub Form_Load()
  Contar
End Sub



Public Sub Converter()
  Dim Tmp1 As String
  Dim Tmp2 As String
  Dim Tmp3 As String
  Dim Tmp4 As Integer
  Dim Tmp5 As Integer
  Dim Tmp6 As String    'Buffer
  Dim Tmp7 As Integer   'Tipo de Procedure
  
  Tmp2 = Dir1.Path
  Select Case Right(Tmp2, 1)
  Case "\"
    Tmp1 = Tmp2 & "Convertido 001"
  Case Else
    Tmp1 = Tmp2 & "\Convertido 001"
    Tmp2 = Tmp2 & "\"
  End Select
  
  Do
    If Dir(Tmp1, vbDirectory) = "" Then Exit Do
    Tmp1 = Left(Tmp1, Len(Tmp1) - 3) & Format(Val(Right(Tmp1, 3)) + 1, "000")
  Loop
  MkDir Tmp1
  
  Tmp1 = Tmp1 & "\"
  
  Tmp3 = Dir(Tmp2 & "*.frm")
  Do Until Tmp3 = ""
    Tmp4 = FreeFile
    Open Tmp2 & Tmp3 For Input As Tmp4
    Tmp5 = FreeFile
    Open Tmp1 & Tmp3 For Output As Tmp5
    
    Tmp7 = 0
    
    Do While Not EOF(Tmp4)
      Line Input #Tmp4, Tmp6
      
      If Left(Tmp6, 48) = "Object = " & Chr(34) & "{6B7E6392-850A-101B-AFC0-4210102A8DA7}" Then
        Tmp6 = "Object = " & Chr(34) & "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0" & Chr(34) & "; " & Chr(34) & "MSCOMCTL.OCX" & Chr(34)
      End If
      
      Rem ---Verificar Begin de objeto--- MSComctlLib
      If Tmp7 = 0 Then
        If Left(Tmp6, 26) = "   Begin ComctlLib.Toolbar" Then
          Tmp7 = 1
          Tmp6 = "   Begin MS" & Mid(Tmp6, 10)
        End If
        If Left(Tmp6, 28) = "   Begin ComctlLib.ImageList" Then
          Tmp7 = 2
          Tmp6 = "   Begin MS" & Mid(Tmp6, 10)
        End If
        If Left(Tmp6, 28) = "   Begin ComctlLib.StatusBar" Then
          Tmp7 = 3
          Tmp6 = "   Begin MS" & Mid(Tmp6, 10)
        End If
        If Left(Tmp6, 25) = "   Begin ComctlLib.Slider" Then
          Tmp7 = 4
          Tmp6 = "   Begin MS" & Mid(Tmp6, 10)
        End If
        If Left(Tmp6, 30) = "   Begin ComctlLib.ProgressBar" Then
          Tmp7 = 5
          Tmp6 = "   Begin MS" & Mid(Tmp6, 10)
        End If
      End If
      Rem ---Verificar End de objeto---
      If Tmp7 <> 0 Then
        If Left(Tmp6, 6) = "   End" Then Tmp7 = 0
      End If
      
      Select Case Tmp7
      Case 1    'Toolbar
        If Tmp6 = "      _Version        =   327682" Then
          Tmp6 = "      _Version        =   393216"
        End If
        If Tmp6 = "      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} " Then
          Tmp6 = "      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} "
        End If
        If Right(Tmp6, 39) = "{0713F354-850A-101B-AFC0-4210102A8DA7} " Then
          Tmp6 = Left(Tmp6, Len(Tmp6) - 39) & "{66833FEA-8583-11D1-B16A-00C0F0283628} "
        End If
      Case 2    'ImageList
        If Tmp6 = "      _Version        =   327682" Then
          Tmp6 = "      _Version        =   393216"
        End If
        If Tmp6 = "      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} " Then
          Tmp6 = "      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} "
        End If
        If Right(Tmp6, 39) = "{0713E8C3-850A-101B-AFC0-4210102A8DA7} " Then
          Tmp6 = Left(Tmp6, Len(Tmp6) - 39) & "{2C247F27-8591-11D1-B16A-00C0F0283628} "
        End If
      Case 3    'StatusBar
        If Tmp6 = "      _Version        =   327682" Then
          Tmp6 = "      _Version        =   393216"
        End If
        If Tmp6 = "      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} " Then
          Tmp6 = "      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} "
        End If
        If Right(Tmp6, 39) = "{0713E89F-850A-101B-AFC0-4210102A8DA7} " Then
          Tmp6 = Left(Tmp6, Len(Tmp6) - 39) & "{8E3867AB-8586-11D1-B16A-00C0F0283628} "
        End If
      Case 4    'Slider
        If Tmp6 = "      _Version        =   327682" Then
          Tmp6 = "      _Version        =   393216"
        End If
      Case 5    'ProgressBar
        If Tmp6 = "      _Version        =   327682" Then
          Tmp6 = "      _Version        =   393216"
        End If
      End Select
      
      Print #Tmp5, Tmp6
    Loop
    
    Close Tmp4
    Close Tmp5
    
    Tmp3 = Dir
  Loop
  
  Tmp3 = Dir(Tmp2 & "*.vbp")
  Do Until Tmp3 = ""
    Tmp4 = FreeFile
    Open Tmp2 & Tmp3 For Input As Tmp4
    Tmp5 = FreeFile
    Open Tmp1 & Tmp3 For Output As Tmp5
    
    Do While Not EOF(Tmp4)
      Line Input #Tmp4, Tmp6
      
      If Left(Tmp6, 45) = "Object={6B7E6392-850A-101B-AFC0-4210102A8DA7}" Then
        Tmp6 = "Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX"
      End If
      Print #Tmp5, Tmp6
    Loop
    
    Close Tmp4
    Close Tmp5
    
    Tmp3 = Dir
  Loop
End Sub
