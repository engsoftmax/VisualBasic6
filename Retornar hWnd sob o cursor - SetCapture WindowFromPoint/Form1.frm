VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINT
  X As Long
  Y As Long
End Type

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex&) As Long
Private Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hWnd&, ByVal lpString$, ByVal cb&)
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Boolean
Private Declare Function WindowFromPoint Lib "user32" (ByVal ptY As Long, ByVal ptX As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex&) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle&, ByVal nWidth&, ByVal crColor&) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject&) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc&, ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&) As Long
Private Declare Sub InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long)
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance&, ByVal lpCursor&) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Const IDC_UPARROW = 32516&
'--------------------------------------------------------------------------
' Module-Level variables
'--------------------------------------------------------------------------

Public mlngHwndCaptured As Long   ' Holds the handle to the captured window

'**************************************************************************
' Purpose:  Turns on SetCapture and changes the mouse pointer when the user
'           clicks down on the form.
'**************************************************************************

Private Sub Form_MouseDown(Button%, Shift%, X As Single, Y As Single)
    If SetCapture(hWnd) Then MousePointer = vbUpArrow
End Sub

'**************************************************************************
' Purpose:  Draws a rectangle around the window currently under the mouse
'           pointer while the primary mouse key is being held down.
'**************************************************************************

Private Sub Form_MouseMove(Button%, Shift%, X As Single, Y As Single)
    Dim pt As POINT          ' Holds the location of the window.
    Static hWndLast As Long  ' The handle of the last window we drew a
                             ' a rectangle on.
    '----------------------------------------------------------------------
    ' If in capture mode, then draw a rectangle around the active window.
    '----------------------------------------------------------------------

    If GetCapture() Then
        '------------------------------------------------------------------
        ' Convert the current mouse position to Screen coordinates.
        '------------------------------------------------------------------

        pt.X = CLng(X)
        pt.Y = CLng(Y)
        ClientToScreen Me.hWnd, pt
        '------------------------------------------------------------------
        ' Pass that value to WindowFromPoint to find out what window we are
        ' pointing to.
        '------------------------------------------------------------------

        mlngHwndCaptured = WindowFromPoint(pt.X, pt.Y)
        '------------------------------------------------------------------
        ' If its not the last window, then erase the previous rectangle
        ' and draw a rectangle around the window under the mouse pointer.
        '------------------------------------------------------------------

        If hWndLast <> mlngHwndCaptured Then
            If hWndLast Then InvertTracker hWndLast
            InvertTracker mlngHwndCaptured
            hWndLast = mlngHwndCaptured
        End If
    End If
  End Sub

'**************************************************************************
' Purpose:  Puts the caption of the window under the cusor into our caption.
'**************************************************************************

Private Sub Form_MouseUp(Button%, Shift%, X As Single, Y As Single)
  Dim strCaption$ ' Buffer used to hold the caption.
  
  '----------------------------------------------------------------------
  ' If a window has been captured, then put its caption in our caption.
  '----------------------------------------------------------------------

  If mlngHwndCaptured Then
    '------------------------------------------------------------------
    ' Create a buffer to hold the caption, and call GetWindowText to
    ' retrive it.
    '------------------------------------------------------------------

    strCaption = Space(1000)
    Caption = Left(strCaption, GetWindowText(mlngHwndCaptured, strCaption, Len(strCaption)))
    '------------------------------------------------------------------
    ' Refresh the entire screen in case we forgot to erase a rectangle.
    '------------------------------------------------------------------

    InvalidateRect 0, 0, True
    '------------------------------------------------------------------
    ' Clear our module-level variable and restore the mouse pointer.
    '------------------------------------------------------------------

    mlngHwndCaptured = False
    MousePointer = vbNormal
  End If
End Sub

'**************************************************************************
' Purpose:  Draws a inverted rectangle around a window on the screen.
' Inputs:   A handle to a enabled and visible window.
'**************************************************************************

Private Sub InvertTracker(hwndDest As Long)
    Dim hdcDest&, hPen&, hOldPen&, hOldBrush&
    Dim cxBorder&, cxFrame&, cyFrame&, cxScreen&, cyScreen&
    Dim rc As RECT, cr As Long
    Const NULL_BRUSH = 5
    Const R2_NOT = 6
    Const PS_INSIDEFRAME = 6
    
    '----------------------------------------------------------------------
    ' Get the screen, border, and frame sizes.
    '----------------------------------------------------------------------

    cxScreen = GetSystemMetrics(0)
    cyScreen = GetSystemMetrics(1)
    cxBorder = GetSystemMetrics(5)
    cxFrame = GetSystemMetrics(32)
    cyFrame = GetSystemMetrics(33)
    '----------------------------------------------------------------------
    ' Get the coordinates of the window on the screen.
    '----------------------------------------------------------------------

    GetWindowRect hwndDest, rc
    '----------------------------------------------------------------------
    ' Get a handle to the window's device context.
    '----------------------------------------------------------------------

    hdcDest = GetWindowDC(hwndDest)
    '----------------------------------------------------------------------
    ' Create an inverse pen that is the size of a window border.
    '----------------------------------------------------------------------

    SetROP2 hdcDest, R2_NOT
    cr = RGB(0, 0, 0)
    hPen = CreatePen(PS_INSIDEFRAME, 3 * cxBorder, cr)
    '----------------------------------------------------------------------
    ' Draw the rectangle around the window.
    '----------------------------------------------------------------------

    hOldPen = SelectObject(hdcDest, hPen)
    hOldBrush = SelectObject(hdcDest, GetStockObject(NULL_BRUSH))
    Rectangle hdcDest, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top
    SelectObject hdcDest, hOldBrush
    SelectObject hdcDest, hOldPen
    '----------------------------------------------------------------------
    ' Give the window its device context back, and destroy our pen.
    '----------------------------------------------------------------------

    ReleaseDC hwndDest, hdcDest
    DeleteObject hPen
End Sub

'**************************************************************************
' Purpose:  Sets up the form, and draws a copy of vbUpArrow on the form.
'**************************************************************************

Private Sub Form_Load()
  '----------------------------------------------------------------------
  ' Size the form and put instructions in the caption.
  '----------------------------------------------------------------------

  Move 0, 0, 250 * Screen.TwipsPerPixelX, 75 * Screen.TwipsPerPixelY
  Caption = "Click & drag the arrow!"
  '----------------------------------------------------------------------
  ' Change the ScaleMode to pixels and turn on AutoRedraw.
  '----------------------------------------------------------------------
 
  ScaleMode = vbPixels
  AutoRedraw = True
  
  Print "Clique e arraste!"
  
  '----------------------------------------------------------------------
  ' Draw vbUpArrow into the form's persistant bitmap.
  '----------------------------------------------------------------------
  
  DrawIcon hdc, (ScaleWidth / 2), 9, LoadCursor(0, IDC_UPARROW)
End Sub



