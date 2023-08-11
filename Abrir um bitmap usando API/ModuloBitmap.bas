Attribute VB_Name = "ModuloBitmap"
Option Explicit

Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Const CBM_INIT = &H4     'initialize bitmap
Public Const SRCCOPY = &HCC0020 'dest = source

Public Const BI_RGB = 0&
Public Const BI_RLE8 = 1&
Public Const BI_RLE4 = 2&
Public Const BI_bitfields = 3&

Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices

Public Type BITMAPINFOHEADER '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Public Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As String * 1024
End Type

Public M_BMinfo As BITMAPINFO
Public M_Bitmap As String     'Bits do Bitmap

Public Function BitMapRowSize(bmwidth As Long, bmbitspixel) As Long
  'Given bitmap width in pixels, and the number of bits
  'per pixel, calculate the bitmap row size in bytes.
  Dim Tmp1 As Long
  
  Tmp1 = bmwidth * bmbitspixel
  
  If (Tmp1 Mod 32) <> 0 Then
    'Cada linha de pixels é armazenada em array de 4 bytes (Long),
    'portanto cada linha deve ser divisivel por 32 bits
    BitMapRowSize = ((Tmp1 + 32 - (Tmp1 Mod 32)) \ 8)
  Else
    BitMapRowSize = Tmp1 \ 8
  End If
End Function

Public Function LerBitmap(ByVal P_Nome As String) As Boolean
  Dim P_NumArq As Integer
  Dim P_FimHead As Long     '+1 = inicio da Paleta ou bitmap RGB
  Dim P_VerBM As String * 2 'Verificar se é bitmap
  Dim P_Paleta As String
  
  LerBitmap = True
  
  'Verificar se o arquivo existe
  If Dir(P_Nome) = "" Then
    LerBitmap = False
    Exit Function
  End If
  
  'Tamanho mínimo para o InfoHead
  If FileLen(P_Nome) < 55 Then
    LerBitmap = False
    Exit Function
  End If
  
  'Abrir e Ler Arquivo
  P_NumArq = FreeFile
  Open P_Nome For Binary As P_NumArq
  
  Get #P_NumArq, 1, P_VerBM
  Get #P_NumArq, 11, P_FimHead
  Get #P_NumArq, 15, M_BMinfo   'Tipo de bitmap
  
  If UCase(P_VerBM) <> "BM" Or M_BMinfo.bmiHeader.biSize <> 40 Or P_FimHead < 54 Then
    LerBitmap = False
    GoTo Fim
  End If
  
  'Paleta de Cores
  If P_FimHead = 54 Then
    M_BMinfo.bmiColors = String(1024, 0)
  Else
    P_Paleta = String(P_FimHead - 54, 0)
    Get #P_NumArq, 55, P_Paleta
    M_BMinfo.bmiColors = P_Paleta & String(1024 - Len(P_Paleta), 0)
  End If
  
  M_Bitmap = String(M_BMinfo.bmiHeader.biSizeImage, 0)
  Get #P_NumArq, P_FimHead + 1, M_Bitmap
  
Fim:
  Close P_NumArq
End Function

Public Sub ExibirBitmapModo1(ByVal P_hDC As Long)
  SetDIBitsToDevice P_hDC, 0, 0, M_BMinfo.bmiHeader.biWidth, M_BMinfo.bmiHeader.biHeight, 0, 0, 0, M_BMinfo.bmiHeader.biHeight, ByVal M_Bitmap, M_BMinfo, DIB_RGB_COLORS
  
  'ou pode ser enviada uma byte array como ByRef
  
  'ReDim P_ByteArray(1 To Len(M_Bitmap)) As Byte
  
  '---------------------------------------------
  'Se for utilizado P_ByteArray = M_Bitmap
  'a array gerada será Unicode(2 bytes por caracter)
  'e não funcionará. Portanto utilize CopyMemory
  'ou
  'For x = 1 To Len(M_Bitmap)
  '  P_ByteArray(x) = Asc(Mid(M_Bitmap, x, 1))
  'Next x
  '---------------------------------------------
  
  'CopyMemory P_ByteArray(1), ByVal M_Bitmap, Len(M_Bitmap)
  
  'SetDIBitsToDevice P_hDC, 0, 0, M_BMinfo.bmiHeader.biWidth, M_BMinfo.bmiHeader.biHeight, 0, 0, 0, M_BMinfo.bmiHeader.biHeight, P_ByteArray(1), M_BMinfo, DIB_RGB_COLORS
End Sub

Public Sub ExibirBitmapModo2(ByVal P_hDC)
  Dim P_hDCmem
  Dim P_hBitmap
  Dim P_hBitmapOld
  
  'Dim P_X As Integer
  'Dim P_Y As Integer
  
  P_hDCmem = CreateCompatibleDC(P_hDC)
  
  If P_hDCmem = 0 Then
    MsgBox "Impossível criar DC handle !", vbExclamation, "Erro"
    Exit Sub
  End If
  
  P_hBitmap = CreateDIBitmap(P_hDC, M_BMinfo.bmiHeader, CBM_INIT, ByVal M_Bitmap, M_BMinfo, DIB_RGB_COLORS)
  
  If P_hBitmap = 0 Then
    DeleteDC P_hDCmem 'Excluir DC criado
    MsgBox "Impossível criar bitmap handle !", vbExclamation, "Erro"
    Exit Sub
  End If
  
  P_hBitmapOld = SelectObject(P_hDCmem, P_hBitmap)
  
  BitBlt P_hDC, 0, 0, M_BMinfo.bmiHeader.biWidth, M_BMinfo.bmiHeader.biHeight, P_hDCmem, 0, 0, SRCCOPY
  
  'ou
  
  'For P_X = 0 To M_BMinfo.bmiHeader.biWidth - 1
  '  For P_Y = 0 To M_BMinfo.bmiHeader.biHeight - 1
  '    SetPixelV P_hDC, P_X, P_Y, GetPixel(P_hDCmem, P_X, P_Y)
  '  Next P_Y
  'Next P_X
  
  SelectObject P_hDCmem, P_hBitmapOld
  DeleteDC P_hDCmem
  DeleteObject P_hBitmap
End Sub

Public Sub ExibirBitmapModo3(ByVal P_hDC)
  Dim P_hDCmem
  Dim P_hBitmap
  Dim P_hBitmapOld
  
  'Dim P_X As Integer
  'Dim P_Y As Integer
  
  P_hDCmem = CreateCompatibleDC(P_hDC)
  
  If P_hDCmem = 0 Then
    MsgBox "Impossível criar DC handle !", vbExclamation, "Erro"
    Exit Sub
  End If
  
  P_hBitmap = CreateCompatibleBitmap(P_hDC, M_BMinfo.bmiHeader.biWidth, M_BMinfo.bmiHeader.biHeight)
  
  If P_hBitmap = 0 Then
    DeleteDC P_hDCmem 'Excluir DC criado
    MsgBox "Impossível criar bitmap handle !", vbExclamation, "Erro"
    Exit Sub
  End If
  
  P_hBitmapOld = SelectObject(P_hDCmem, P_hBitmap)
  
  SetDIBits P_hDCmem, P_hBitmap, 0, M_BMinfo.bmiHeader.biHeight, ByVal M_Bitmap, M_BMinfo, DIB_RGB_COLORS
  
  BitBlt P_hDC, 0, 0, M_BMinfo.bmiHeader.biWidth, M_BMinfo.bmiHeader.biHeight, P_hDCmem, 0, 0, SRCCOPY
  
  'ou
  
  'For P_X = 0 To M_BMinfo.bmiHeader.biWidth - 1
  '  For P_Y = 0 To M_BMinfo.bmiHeader.biHeight - 1
  '    SetPixelV P_hDC, P_X, P_Y, GetPixel(P_hDCmem, P_X, P_Y)
  '  Next P_Y
  'Next P_X
  
  SelectObject P_hDCmem, P_hBitmapOld
  DeleteDC P_hDCmem
  DeleteObject P_hBitmap
End Sub
