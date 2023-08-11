VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menu Iniciar e impressora"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003

Private Const KEY_ALL_ACCESS As Long = &HF0063
Private Const ERROR_SUCCESS As Long = 0
Private Const REG_SZ As Long = 1
Public Function VerRegString(P_hInKey As Long, ByVal P_SubKey, ByVal P_ValName) As String
  Dim P_RetVal As String
  Dim P_hSubKey As Long
  Dim P_dwType As Long
  Dim P_Retorno As Long
  Dim P_Valor As String
  
  P_RetVal = ""
  P_Retorno = RegOpenKeyEx(P_hInKey, P_SubKey, 0&, KEY_ALL_ACCESS, P_hSubKey)
  If P_Retorno <> ERROR_SUCCESS Then GoTo Quit_Now
  
  P_Valor = String(256, 0)
  
  P_Retorno = RegQueryValueEx(P_hSubKey, P_ValName, 0&, P_dwType, ByVal P_Valor, 256)
  If P_Retorno = ERROR_SUCCESS And P_dwType = REG_SZ Then
    P_RetVal = Left(P_Valor, 256)
  Else
    P_RetVal = "--Not String--"
  End If
  
  If P_hInKey = 0 Then P_Retorno = RegCloseKey(P_hSubKey)
  
  If InStr(1, P_RetVal, Chr(0)) Then
    P_RetVal = Left(P_RetVal, InStr(1, P_RetVal, Chr(0)) - 1)
  End If
  
Quit_Now:
  
  VerRegString = P_RetVal
End Function

Private Sub Form_Load()
  AutoRedraw = True
  Show
  
  Print "Impressora Padrão:"
  Print vbTab; VerRegString(HKEY_CURRENT_CONFIG, "System\CurrentControlSet\Control\Print\Printers", "Default")
  
  Print
  
  Print "Menu Iniciar:"
  Print vbTab; VerRegString(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Start Menu")
  
  Print
  
  Print "Pasta Programas do menu Iniciar:"
  Print vbTab; VerRegString(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs")
  
  Print
  
  Print "Pasta Iniciar do menu Programas:"
  Print vbTab; VerRegString(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
  
  Print
  
  Print "Papel de Parede:"
  Print vbTab; VerRegString(HKEY_CURRENT_USER, "Control Panel\desktop", "WallPaper")
End Sub


