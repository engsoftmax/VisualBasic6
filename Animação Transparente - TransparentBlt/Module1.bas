Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Type BITMAP '14 bytes
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Public Const SRCCOPY = &HCC0020     ' (DWORD) dest = source
Public Const SRCAND = &H8800C6      ' (DWORD) dest = source AND dest
Public Const SRCPAINT = &HEE0086    ' (DWORD) dest = source OR dest
Public Const NOTSRCCOPY = &H330008  ' (DWORD) dest = (NOT source)

Sub TransparentBlt(dest As Control, ByVal srcBmp As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal TransColor As Long)
      Const PIXEL = 3
      
      Dim destScale As Long
      Dim srcDC As Long             'source bitmap (color)
      Dim saveDC As Long            'backup copy of source bitmap
      Dim maskDC As Long            'mask bitmap (monochrome)
      Dim invDC As Long             'inverse of mask bitmap (monochrome)
      Dim resultDC As Long          'combination of source bitmap & background
      Dim bmp As BITMAP             'description of the source bitmap
      Dim hResultBmp As Long        'Bitmap combination of source & background
      Dim hSaveBmp As Long          'Bitmap stores backup copy of source bitmap
      Dim hMaskBmp As Long          'Bitmap stores mask (monochrome)
      Dim hInvBmp As Long           'Bitmap holds inverse of mask (monochrome)
      Dim hPrevBmp As Long          'Bitmap holds previous bitmap selected in DC
      Dim hSrcPrevBmp As Long       'Holds previous bitmap in source DC
      Dim hSavePrevBmp As Long      'Holds previous bitmap in saved DC
      Dim hDestPrevBmp As Long      'Holds previous bitmap in destination DC
      Dim hMaskPrevBmp As Long      'Holds previous bitmap in the mask DC
      Dim hInvPrevBmp As Long       'Holds previous bitmap in inverted mask DC
      Dim OrigColor As Long         'Holds original background color from source DC
      Dim Success As Long           'Stores result of call to Windows API
      
      If TypeOf dest Is PictureBox Then 'Ensure objects are picture boxes
        destScale = dest.ScaleMode  'Store ScaleMode to restore later
        dest.ScaleMode = PIXEL      'Set ScaleMode to pixels for Windows GDI
        
        'Retrieve bitmap to get width (bmp.bmWidth) & height (bmp.bmHeight)
        Success = GetObject(srcBmp, Len(bmp), bmp)
        srcDC = CreateCompatibleDC(dest.hdc)    'Create DC to hold stage
        saveDC = CreateCompatibleDC(dest.hdc)   'Create DC to hold stage
        maskDC = CreateCompatibleDC(dest.hdc)   'Create DC to hold stage
        invDC = CreateCompatibleDC(dest.hdc)    'Create DC to hold stage
        resultDC = CreateCompatibleDC(dest.hdc) 'Create DC to hold stage
        
        'Create monochrome bitmaps for the mask-related bitmaps:
        hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
        hInvBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
        
        'Create color bitmaps for final result & stored copy of source
        hResultBmp = CreateCompatibleBitmap(dest.hdc, bmp.bmWidth, bmp.bmHeight)
        hSaveBmp = CreateCompatibleBitmap(dest.hdc, bmp.bmWidth, bmp.bmHeight)
        
        hSrcPrevBmp = SelectObject(srcDC, srcBmp)     'Select bitmap in DC
        hSavePrevBmp = SelectObject(saveDC, hSaveBmp) 'Select bitmap in DC
        hMaskPrevBmp = SelectObject(maskDC, hMaskBmp) 'Select bitmap in DC
        hInvPrevBmp = SelectObject(invDC, hInvBmp)    'Select bitmap in DC
        hDestPrevBmp = SelectObject(resultDC, hResultBmp) 'Select bitmap
        
        Success = BitBlt(saveDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCCOPY)           'Make backup of source bitmap to restore later
        
        'Create mask: set background color of source to transparent color.
        OrigColor = SetBkColor(srcDC, TransColor)
        Success = BitBlt(maskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCCOPY)
        TransColor = SetBkColor(srcDC, OrigColor)
        
        'Create inverse of mask to AND w/ source & combine w/ background.
        Success = BitBlt(invDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, NOTSRCCOPY)
        
        'Copy background bitmap to result & create final transparent bitmap
        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, dest.hdc, destX, destY, SRCCOPY)
        
        'AND mask bitmap w/ result DC to punch hole in the background by
        'painting black area for non-transparent portion of source bitmap.
        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, SRCAND)
        
        'AND inverse mask w/ source bitmap to turn off bits associated
        'with transparent area of source bitmap by making it black.
        Success = BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, invDC, 0, 0, SRCAND)
        
        'XOR result w/ source bitmap to make background show through.
        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCPAINT)
        Success = BitBlt(dest.hdc, destX, destY, bmp.bmWidth, bmp.bmHeight, resultDC, 0, 0, SRCCOPY)           'Display transparent bitmap on backgrnd
        Success = BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, saveDC, 0, 0, SRCCOPY)           'Restore backup of bitmap.
        
        hPrevBmp = SelectObject(srcDC, hSrcPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(saveDC, hSavePrevBmp) 'Select orig object
        hPrevBmp = SelectObject(resultDC, hDestPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(maskDC, hMaskPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(invDC, hInvPrevBmp) 'Select orig object
        
        Success = DeleteObject(hSaveBmp)   'Deallocate system resources.
        Success = DeleteObject(hMaskBmp)   'Deallocate system resources.
        Success = DeleteObject(hInvBmp)    'Deallocate system resources.
        Success = DeleteObject(hResultBmp) 'Deallocate system resources.
        Success = DeleteDC(srcDC)          'Deallocate system resources.
        Success = DeleteDC(saveDC)         'Deallocate system resources.
        Success = DeleteDC(invDC)          'Deallocate system resources.
        Success = DeleteDC(maskDC)         'Deallocate system resources.
        Success = DeleteDC(resultDC)       'Deallocate system resources.
        
        dest.ScaleMode = destScale 'Restore ScaleMode of destination.
    End If
End Sub
