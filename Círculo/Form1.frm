VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Círculo"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2190
   LinkTopic       =   "Form1"
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   146
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Tmp1 As Integer
  Dim Tmp2 As Single
  Dim Tmp3 As Integer
  Dim Tmp4 As Integer
  
  Cls
  
  Line (200, 0)-(200, 400), 0
  Line (0, 200)-(400, 200), 0
  
  Tmp3 = Abs(X - 200)
  Tmp4 = Abs(Y - 200)
  
  If Tmp3 > Tmp4 Then Tmp1 = Tmp3 Else Tmp1 = Tmp4
  If Tmp3 > 0 Then
    Tmp2 = Tmp4 / Tmp3
  Else
    Tmp2 = 10 ^ 10  'Evita divisão por zero
  End If
  
  Circle (200, 200), Tmp1, QBColor(12), , , Tmp2
  
  Print
  Print " Aspecto: "; CDec(Tmp2)
  Print
  Print " Raio: "; Tmp1
End Sub



