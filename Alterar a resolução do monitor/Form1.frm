VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alterar Resolução"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   202
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "800 x 600"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Área de trabalho:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32

Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const DM_BITSPERPEL = &H40000

Private Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

Dim DevM() As DEVMODE

Private Sub Command1_Click()
  Dim Tmp1 As Long
  
  If Combo1.ListIndex < 0 Then Exit Sub
  
  DevM(Combo1.ItemData(Combo1.ListIndex)).dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
  
  Tmp1 = ChangeDisplaySettings(DevM(Combo1.ItemData(Combo1.ListIndex)), 0&)
End Sub


Private Sub Command2_Click()
  Dim Tmp1 As DEVMODE
  Dim Tmp2 As Boolean
  
  Tmp2 = EnumDisplaySettings(0&, Combo1.ItemData(Combo1.ListCount - 1), Tmp1)
  
  If Tmp2 = False Then
    Beep
    Exit Sub
  End If
  
  Tmp1.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
  
  Tmp1.dmPelsWidth = 800
  Tmp1.dmPelsHeight = 600
  
  ChangeDisplaySettings Tmp1, 0&
End Sub


Private Sub Form_Load()
  InicializarDevM
  
  If Combo1.ListCount = 0 Then
    Command1.Enabled = False
    Command2.Enabled = False
  Else
    Combo1.ListIndex = 0
  End If
End Sub



Public Sub InicializarDevM()
  Dim Tmp1 As Boolean
  Dim Tmp2 As Integer
  
  Tmp2 = 0
  Do
    ReDim Preserve DevM(0 To Tmp2)
    Tmp1 = EnumDisplaySettings(0&, Tmp2, DevM(Tmp2))
    
    If Tmp1 Then
      Combo1.AddItem DevM(Tmp2).dmPelsWidth & " x " & DevM(Tmp2).dmPelsHeight & " x " & DevM(Tmp2).dmBitsPerPel & " bits"
      Combo1.ItemData(Combo1.NewIndex) = Tmp2
    End If
    
    Tmp2 = Tmp2 + 1
  Loop Until (Tmp1 = False)
End Sub
