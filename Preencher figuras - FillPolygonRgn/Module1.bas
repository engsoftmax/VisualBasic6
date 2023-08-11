Attribute VB_Name = "Module1"
Option Explicit

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const ALTERNATE = 1   'ALTERNATE and WINDING are
Public Const WINDING = 2     'constants for FillMode.
Public Const BLACKBRUSH = 4  'Constant for brush type.
