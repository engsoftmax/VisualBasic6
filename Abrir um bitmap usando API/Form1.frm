VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
  Dim P_Retorno As Boolean
  Dim P_MSG As String
  
  CommonDialog1.Filter = "Bitmap do Windows (*.bmp)|*.bmp"
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
  CommonDialog1.DialogTitle = "Abrir Figura"
  CommonDialog1.FileName = ""
  CommonDialog1.ShowOpen
  
  If CommonDialog1.FileName = "" Then Exit Sub
  
  Picture = Nothing
  
  P_Retorno = LerBitmap(CommonDialog1.FileName)
  
  If Not P_Retorno Then
    MsgBox CommonDialog1.FileName & vbCr & vbCr & "Este não é um arquivo de bitmap válido!", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  Width = (8 + M_BMinfo.bmiHeader.biWidth) * Screen.TwipsPerPixelX
  Height = (28 + M_BMinfo.bmiHeader.biHeight) * Screen.TwipsPerPixelY
  
  Refresh
  
  ExibirBitmapModo1 hdc
  
  M_Bitmap = "" 'Liberar memória
  
  Refresh
  
  'Informações do bitmap
  
  P_MSG = P_MSG & M_BMinfo.bmiHeader.biWidth & " x " & M_BMinfo.bmiHeader.biHeight & " pixels" & vbCr & vbCr
  
  Select Case M_BMinfo.bmiHeader.biBitCount
  Case 1
    P_MSG = P_MSG & "Monocromático" & vbCr & vbCr
  Case 4
    P_MSG = P_MSG & "16 cores" & vbCr & vbCr
  Case 8
    P_MSG = P_MSG & "256 cores" & vbCr & vbCr
  Case Else
    P_MSG = P_MSG & M_BMinfo.bmiHeader.biBitCount & " bits" & vbCr & vbCr
  End Select
  
  Select Case M_BMinfo.bmiHeader.biCompression
  Case BI_RLE4, BI_RLE8
    P_MSG = P_MSG & "Compactação RLE"
  Case BI_RGB, BI_bitfields
    P_MSG = P_MSG & "Não Compactado"
  End Select
  
  MsgBox P_MSG, vbInformation, "Tipo de Bitmap"
End Sub

Private Sub Form_Load()
  Font.Name = "Times New Roman"
  Font.Size = 18
  Font.Bold = True
  
  ForeColor = QBColor(8)
  CurrentX = 66
  CurrentY = 67
  Print "Clique aqui!"
  
  ForeColor = QBColor(0)
  CurrentX = 64
  CurrentY = 64
  Print "Clique aqui!"
End Sub


