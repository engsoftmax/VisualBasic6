Attribute VB_Name = "ModuloParent"
Option Explicit

Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Public Const HWND_BOTTOM = 1&
Public Const HWND_TOP = 0&
Public Const HWND_TOPMOST = -1&
Public Const HWND_NOTOPMOST = -2&

Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Sub Main()
  Dim Tmp1 As Long      'DisplayHwnd
  Dim Tmp2 As RECT      'ClientRect
  Dim Tmp3 As Long
  
  Tmp1 = Val(InputBox("Informe o hWnd:", "Parent"))
  If Tmp1 = 0 Then End
  
  Rem ---Buscar ClientRect---
  GetClientRect Tmp1, Tmp2
  
  Load frmParent
  
  Rem ---Get current window style---
  Tmp3 = GetWindowLong(frmParent.hwnd, GWL_STYLE)
  Rem ---Append "WS_CHILD" style to the hWnd window style---
  Tmp3 = Tmp3 Or WS_CHILD
  Rem ---Add new style to window---
  SetWindowLong frmParent.hwnd, GWL_STYLE, Tmp3
  
  Rem ---Set preview window as parent window---
  SetParent frmParent.hwnd, Tmp1
  Rem ---Save the hWnd Parent in hWnd's window struct---
  SetWindowLong frmParent.hwnd, GWL_HWNDPARENT, Tmp1
  
  Rem ---Show screensaver in the preview window---
  frmParent.Show
  'SetWindowPos frmParent.hwnd, HWND_TOP, 0&, 0&, Tmp2.Right, Tmp2.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
