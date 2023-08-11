VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Configurações regionais"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Formato da data abreviada:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1950
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_SETTINGCHANGE = &H1A
Private Const HWND_BROADCAST = &HFFFF&

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Locale Types.
'These types are used for the GetLocaleInfoW NLS API routine.

'LOCALE_NOUSEROVERRIDE is also used in GetTimeFormatW and GetDateFormatW.
Private Const LOCALE_NOUSEROVERRIDE = &H80000000  'do not use user overrides

Private Const LOCALE_ILANGUAGE = &H1         'language id
Private Const LOCALE_SLANGUAGE = &H2         'localized name of language
Private Const LOCALE_SENGLANGUAGE = &H1001   'English name of language
Private Const LOCALE_SABBREVLANGNAME = &H3   'abbreviated language name
Private Const LOCALE_SNATIVELANGNAME = &H4   'native name of language
Private Const LOCALE_ICOUNTRY = &H5          'country code
Private Const LOCALE_SCOUNTRY = &H6          'localized name of country
Private Const LOCALE_SENGCOUNTRY = &H1002    'English name of country
Private Const LOCALE_SABBREVCTRYNAME = &H7   'abbreviated country name
Private Const LOCALE_SNATIVECTRYNAME = &H8   'native name of country
Private Const LOCALE_IDEFAULTLANGUAGE = &H9  'default language id
Private Const LOCALE_IDEFAULTCOUNTRY = &HA   'default country code
Private Const LOCALE_IDEFAULTCODEPAGE = &HB  'default code page

Private Const LOCALE_SLIST = &HC             'list item separator
Private Const LOCALE_IMEASURE = &HD          '0 = metric, 1 = US

Private Const LOCALE_SDECIMAL = &HE          'decimal separator
Private Const LOCALE_STHOUSAND = &HF         'thousand separator
Private Const LOCALE_SGROUPING = &H10        'digit grouping
Private Const LOCALE_IDIGITS = &H11          'number of fractional digits
Private Const LOCALE_ILZERO = &H12           'leading zeros for decimal
Private Const LOCALE_SNATIVEDIGITS = &H13    'native ascii 0-9

Private Const LOCALE_SCURRENCY = &H14        'local monetary symbol
Private Const LOCALE_SINTLSYMBOL = &H15      'intl monetary symbol
Private Const LOCALE_SMONDECIMALSEP = &H16   'monetary decimal separator
Private Const LOCALE_SMONTHOUSANDSEP = &H17  'monetary thousand separator
Private Const LOCALE_SMONGROUPING = &H18     'monetary grouping
Private Const LOCALE_ICURRDIGITS = &H19      '# local monetary digits
Private Const LOCALE_IINTLCURRDIGITS = &H1A  '# intl monetary digits
Private Const LOCALE_ICURRENCY = &H1B        'positive currency mode
Private Const LOCALE_INEGCURR = &H1C         'negative currency mode

Private Const LOCALE_SDATE = &H1D            'date separator
Private Const LOCALE_STIME = &H1E            'time separator
Private Const LOCALE_SSHORTDATE = &H1F       'short date format string
Private Const LOCALE_SLONGDATE = &H20        'long date format string
Private Const LOCALE_STIMEFORMAT = &H1003    'time format string
Private Const LOCALE_IDATE = &H21            'short date format ordering
Private Const LOCALE_ILDATE = &H22           'long date format ordering
Private Const LOCALE_ITIME = &H23            'time format specifier
Private Const LOCALE_ICENTURY = &H24         'century format specifier
Private Const LOCALE_ITLZERO = &H25          'leading zeros in time field
Private Const LOCALE_IDAYLZERO = &H26        'leading zeros in day field
Private Const LOCALE_IMONLZERO = &H27        'leading zeros in month field
Private Const LOCALE_S1159 = &H28            'AM designator
Private Const LOCALE_S2359 = &H29            'PM designator

Private Const LOCALE_SDAYNAME1 = &H2A        'long name for Monday
Private Const LOCALE_SDAYNAME2 = &H2B        'long name for Tuesday
Private Const LOCALE_SDAYNAME3 = &H2C        'long name for Wednesday
Private Const LOCALE_SDAYNAME4 = &H2D        'long name for Thursday
Private Const LOCALE_SDAYNAME5 = &H2E        'long name for Friday
Private Const LOCALE_SDAYNAME6 = &H2F        'long name for Saturday
Private Const LOCALE_SDAYNAME7 = &H30        'long name for Sunday
Private Const LOCALE_SABBREVDAYNAME1 = &H31  'abbreviated name for Monday
Private Const LOCALE_SABBREVDAYNAME2 = &H32  'abbreviated name for Tuesday
Private Const LOCALE_SABBREVDAYNAME3 = &H33  'abbreviated name for Wednesday
Private Const LOCALE_SABBREVDAYNAME4 = &H34  'abbreviated name for Thursday
Private Const LOCALE_SABBREVDAYNAME5 = &H35  'abbreviated name for Friday
Private Const LOCALE_SABBREVDAYNAME6 = &H36  'abbreviated name for Saturday
Private Const LOCALE_SABBREVDAYNAME7 = &H37  'abbreviated name for Sunday
Private Const LOCALE_SMONTHNAME1 = &H38      'long name for January
Private Const LOCALE_SMONTHNAME2 = &H39      'long name for February
Private Const LOCALE_SMONTHNAME3 = &H3A      'long name for March
Private Const LOCALE_SMONTHNAME4 = &H3B      'long name for April
Private Const LOCALE_SMONTHNAME5 = &H3C      'long name for May
Private Const LOCALE_SMONTHNAME6 = &H3D      'long name for June
Private Const LOCALE_SMONTHNAME7 = &H3E      'long name for July
Private Const LOCALE_SMONTHNAME8 = &H3F      'long name for August
Private Const LOCALE_SMONTHNAME9 = &H40      'long name for September
Private Const LOCALE_SMONTHNAME10 = &H41     'long name for October
Private Const LOCALE_SMONTHNAME11 = &H42     'long name for November
Private Const LOCALE_SMONTHNAME12 = &H43     'long name for December
Private Const LOCALE_SABBREVMONTHNAME1 = &H44 'abbreviated name for January
Private Const LOCALE_SABBREVMONTHNAME2 = &H45 'abbreviated name for February
Private Const LOCALE_SABBREVMONTHNAME3 = &H46 'abbreviated name for March
Private Const LOCALE_SABBREVMONTHNAME4 = &H47 'abbreviated name for April
Private Const LOCALE_SABBREVMONTHNAME5 = &H48 'abbreviated name for May
Private Const LOCALE_SABBREVMONTHNAME6 = &H49 'abbreviated name for June
Private Const LOCALE_SABBREVMONTHNAME7 = &H4A 'abbreviated name for July
Private Const LOCALE_SABBREVMONTHNAME8 = &H4B 'abbreviated name for August
Private Const LOCALE_SABBREVMONTHNAME9 = &H4C 'abbreviated name for September
Private Const LOCALE_SABBREVMONTHNAME10 = &H4D 'abbreviated name for October
Private Const LOCALE_SABBREVMONTHNAME11 = &H4E 'abbreviated name for November
Private Const LOCALE_SABBREVMONTHNAME12 = &H4F 'abbreviated name for December
Private Const LOCALE_SABBREVMONTHNAME13 = &H100F

Private Const LOCALE_SPOSITIVESIGN = &H50    'positive sign
Private Const LOCALE_SNEGATIVESIGN = &H51    'negative sign
Private Const LOCALE_IPOSSIGNPOSN = &H52     'positive sign position
Private Const LOCALE_INEGSIGNPOSN = &H53     'negative sign position
Private Const LOCALE_IPOSSYMPRECEDES = &H54  'mon sym precedes pos amt
Private Const LOCALE_IPOSSEPBYSPACE = &H55   'mon sym sep by space from pos amt
Private Const LOCALE_INEGSYMPRECEDES = &H56  'mon sym precedes neg amt
Private Const LOCALE_INEGSEPBYSPACE = &H57   'mon sym sep by space from neg amt

'Time Flags for GetTimeFormatW.
Private Const TIME_NOMINUTESORSECONDS = &H1  'do not use minutes or seconds
Private Const TIME_NOSECONDS = &H2           'do not use seconds
Private Const TIME_NOTIMEMARKER = &H4        'do not use time marker
Private Const TIME_FORCE24HOURFORMAT = &H8   'always use 24 hour format

'Date Flags for GetDateFormatW.
Private Const DATE_SHORTDATE = &H1           'use short date picture
Private Const DATE_LONGDATE = &H2            'use long date picture

'Code Page Dependent APIs
Private Const MAX_DEFAULTCHAR = 2
Private Const MAX_LEADBYTES = 12             '5 ranges, 2 bytes ea., 0 term.

Private Type CPINFO
  MaxCharSize As Long                        'max length (Byte) of a char
  DefaultChar(MAX_DEFAULTCHAR) As Byte       'default character
  LeadByte(MAX_LEADBYTES) As Byte            'lead byte ranges
End Type

Private Declare Function IsValidCodePage Lib "kernel32" (ByVal CodePage As Long) As Long
Private Declare Function GetACP Lib "kernel32" () As Long
Private Declare Function GetOEMCP Lib "kernel32" () As Long
Private Declare Function GetCPInfo Lib "kernel32" (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
Private Declare Function IsDBCSLeadByte Lib "kernel32" (ByVal bTestChar As Byte) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

'Locale Dependent APIs
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type


Private Declare Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Private Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Private Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

'Locale Independent APIs
Private Declare Function GetStringTypeA Lib "kernel32" (ByVal lcid As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Long) As Long
Private Declare Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long


Private Sub Command1_Click()
  Dim P_LCID As Long
  
  If SetLocaleInfo(P_LCID, LOCALE_SSHORTDATE, Text1.Text) = 0 Then
    MsgBox "Não foi possível completar a operação!", vbInformation, "Erro"
    Exit Sub
  End If
  
  PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub

Private Sub Form_Load()
  Dim P_Buffer As String
  Dim P_Return As Long
  Dim P_LCID As Long
  
  P_Buffer = String(255, 0)
  P_LCID = GetSystemDefaultLCID()
  
  P_Return = GetLocaleInfo(P_LCID, LOCALE_SSHORTDATE, P_Buffer, 255)
  Text1.Text = Left(P_Buffer, P_Return)
End Sub


