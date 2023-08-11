VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fill Polygon RGN"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Clique aqui"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim NumCoords As Long
  Dim bool As Long
  Dim hBrush As Long
  Dim hRgn As Long

  'Dimension coordinate array.
  ReDim Poly(1 To 10) As POINTAPI
  
  'Set scalemode to pixels to set up points of triangle.
  Form1.ScaleMode = 3
  
  'Assign values to points.
  Poly(1).x = 10
  Poly(1).y = 10
  Poly(2).x = 80
  Poly(2).y = 10
  Poly(3).x = 80
  Poly(3).y = 30
  Poly(4).x = 30
  Poly(4).y = 30
  Poly(5).x = 30
  Poly(5).y = 50
  Poly(6).x = 80
  Poly(6).y = 50
  Poly(7).x = 80
  Poly(7).y = 70
  Poly(8).x = 30
  Poly(8).y = 70
  Poly(9).x = 30
  Poly(9).y = 120
  Poly(10).x = 10
  Poly(10).y = 120
  
  'Number of vertices in polygon.
  NumCoords = 10
  
  'Polygon function creates unfilled polygon on screen.
  
  'Remark FillRgn statement to see results.
  Polygon Form1.hdc, Poly(1), NumCoords
  
  'Gets stock black brush.
  hBrush = GetStockObject(BLACKBRUSH)
  
  'Creates region to fill with color.
  hRgn = CreatePolygonRgn(Poly(1), NumCoords, ALTERNATE)
  
  'If the creation of the region was successful then color.
  If hRgn Then
    FillRgn Form1.hdc, hRgn, hBrush
  End If
  
  'Liberar hRgn
  DeleteObject hRgn
End Sub

