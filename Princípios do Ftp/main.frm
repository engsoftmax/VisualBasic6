VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "FTP incompleto"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox filLocal 
      DragIcon        =   "main.frx":0000
      Height          =   1845
      Hidden          =   -1  'True
      Left            =   120
      MultiSelect     =   2  'Extended
      System          =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   3135
   End
   Begin VB.DirListBox dirLocal 
      Height          =   1440
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.DriveListBox drvLocal 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin ComctlLib.StatusBar sbFTP 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5070
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9340
            TextSave        =   ""
            Key             =   "status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   847
            MinWidth        =   847
            TextSave        =   ""
            Key             =   "action"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   847
            MinWidth        =   847
            Picture         =   "main.frx":0442
            TextSave        =   ""
            Key             =   "connect"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox lstRemoteFiles 
      DragIcon        =   "main.frx":075C
      Height          =   3375
      Left            =   3360
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblUserID 
      Alignment       =   1  'Right Justify
      Caption         =   "UserID:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblURL 
      Alignment       =   1  'Right Justify
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgDisconnected 
      Height          =   480
      Left            =   6480
      Picture         =   "main.frx":0B9E
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgConnected 
      Height          =   480
      Left            =   6480
      Picture         =   "main.frx":0FE0
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ftpDIR    As Integer = 0
Private Const ftpPUT    As Integer = 1
Private Const ftpGET    As Integer = 2
Private Const ftpDEL    As Integer = 3

Private iLastFTP        As Integer
Private Sub cmdConnect_Click()
    On Error GoTo ConnectError
    Inet1.URL = txtURL
    Inet1.UserName = txtUserName
    Inet1.Password = txtPassword
    Inet1.Protocol = icFTP
    iLastFTP = ftpDIR
    
    Inet1.Execute Inet1.URL, "DIR"
    
    Exit Sub
ConnectError:
    sbFTP.Panels("status").Text = Err.Description
End Sub


Private Sub dirLocal_Change()
    filLocal.Path = dirLocal.Path
End Sub

Private Sub drvLocal_Change()
    dirLocal.Path = drvLocal.Drive
End Sub


Private Sub filLocal_DragDrop(Source As Control, X As Single, Y As Single)
    'receiving files from FTP site.
    Dim i           As Integer
    Dim sFileList   As String
    
    If TypeOf Source Is ListBox Then
        For i = 0 To Source.ListCount - 1
            If Source.Selected(i) Then
                sFileList = sFileList & Source.List(i) & "|"
            End If
        Next
    End If
    If Len(sFileList) > 0 Then
        'strip off the last pipe
        sFileList = Left(sFileList, Len(sFileList) - 1)
        GetFiles sFileList
    End If
    
End Sub

Private Sub filLocal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    filLocal.Drag vbBeginDrag
    
End Sub

Private Sub filLocal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    filLocal.Drag vbEndDrag
End Sub


Private Sub Inet1_StateChanged(ByVal State As Integer)

    Select Case State
        Case icNone
            sbFTP.Panels("status").Text = ""
        Case icResolvingHost
            sbFTP.Panels("status").Text = "Resolving Host"
        Case icHostResolved
            sbFTP.Panels("status").Text = "Host Resolved"
        Case icConnecting
            sbFTP.Panels("status").Text = "Connecting..."
        Case icConnected
            sbFTP.Panels("status").Text = "Connected!"
        Case icRequesting
            sbFTP.Panels("status").Text = "Requesting..."
        Case icRequestSent
            sbFTP.Panels("status").Text = "Request Sent"
        Case icReceivingResponse
            sbFTP.Panels("status").Text = "Receiving Response..."
        Case icResponseReceived
            sbFTP.Panels("status").Text = "Response Received!"
        Case icDisconnecting
            sbFTP.Panels("status").Text = "Disconnecting..."
        Case icDisconnected
            sbFTP.Panels("status").Text = "Disconnected"
        Case icError
            sbFTP.Panels("status").Text = "Error! " & Trim(CStr(Inet1.ResponseCode)) & ": " & Inet1.ResponseInfo
        Case icResponseCompleted
            sbFTP.Panels("status").Text = "Response Completed!"
            ReactToResponse iLastFTP
    End Select
    Debug.Print Inet1.ResponseInfo & " -- " & sbFTP.Panels("status").Text
End Sub



Public Function ReactToResponse(ByVal iLastCommand As Integer) As Long
    Select Case iLastCommand
        Case ftpDIR
            ShowRemoteFileList
        Case ftpPUT
            Debug.Print "*** File Transferred (Sent)"
            MsgBox "File Sent from directory " & CurDir()
        Case ftpGET
            Debug.Print "*** File Transferred (Received)"
            MsgBox "File Received and placed in directory " & CurDir()
        Case ftpDEL
    End Select

End Function

Public Function ShowRemoteFileList() As Long

    Dim sFileList       As String
    Dim sTemp           As String
    Dim p               As Integer
    
    sTemp = Inet1.GetChunk(1024)
    Do While Len(sTemp) > 0
        DoEvents
        sFileList = sFileList & sTemp
        sTemp = Inet1.GetChunk(1024)
    Loop
    
    lstRemoteFiles.Clear
    Do While sFileList > ""
        DoEvents
        p = InStr(sFileList, vbCrLf)
        If p > 0 Then
            lstRemoteFiles.AddItem Left(sFileList, p - 1)
            If Len(sFileList) > (p + 2) Then
                sFileList = Mid(sFileList, p + 2)
            Else
                sFileList = ""
            End If
        Else
            lstRemoteFiles.AddItem sFileList
            sFileList = ""
        End If
    Loop
End Function

Private Sub lstRemoteFiles_DblClick()
    Dim sSelText    As String
    sSelText = lstRemoteFiles.List(lstRemoteFiles.ListIndex)
    If sSelText = "../" Or sSelText = "..\" Then
        'send cd..
        'send dir
    End If
End Sub


Private Sub lstRemoteFiles_DragDrop(Source As Control, X As Single, Y As Single)
    Dim i           As Integer
    Dim sFileList   As String
    
    If TypeOf Source Is FileListBox Then
        For i = 0 To Source.ListCount - 1
            If Source.Selected(i) Then
                sFileList = sFileList & Source.List(i) & "|"
            End If
        Next
    End If
    If Len(sFileList) > 0 Then
        'strip off the last pipe
        sFileList = Left(sFileList, Len(sFileList) - 1)
        PutFiles sFileList
    End If
End Sub



Public Function PutFiles(sFileList As String) As Long
    
    Dim sFile       As String
    Dim sTemp       As String
    Dim p           As Integer
    
    iLastFTP = ftpPUT
    sTemp = sFileList
    Do While sTemp > ""
        DoEvents
        p = InStr(sTemp, "|")
        If p Then
            sFile = Left(sTemp, p - 1)
            sTemp = Mid(sTemp, p + 1)
        Else
            sFile = sTemp
            sTemp = ""
        End If
        Debug.Print "PUT " & sFile & " " & sFile
        Inet1.Execute Inet1.URL, "PUT " & sFile & " " & sFile
        
        'wait until this execution is done before going to next file
        Do
            DoEvents
        Loop Until Not Inet1.StillExecuting
    Loop
    
    iLastFTP = ftpDIR
    Inet1.Execute Inet1.URL, "DIR"
        
End Function

Private Sub lstRemoteFiles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Inet1.Execute Inet1.URL, "DEL " & lstRemoteFiles.List(lstRemoteFiles.ListIndex)
        Do
            DoEvents
        Loop While Inet1.StillExecuting
    End If
    iLastFTP = ftpDIR
    Inet1.Execute Inet1.URL, "DIR"
End Sub

Private Sub lstRemoteFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstRemoteFiles.Drag vbBeginDrag
End Sub


Private Sub lstRemoteFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstRemoteFiles.Drag vbEndDrag
End Sub



Public Function GetFiles(sFileList As String) As Long
    
    Dim sFile       As String
    Dim sTemp       As String
    Dim p           As Integer
    
    iLastFTP = ftpGET
    sTemp = sFileList
    Do While sTemp > ""
        DoEvents
        p = InStr(sTemp, "|")
        If p Then
            sFile = Left(sTemp, p - 1)
            sTemp = Mid(sTemp, p + 1)
        Else
            sFile = sTemp
            sTemp = ""
        End If
        Debug.Print "GET " & sFile & " " & sFile
        Inet1.Execute Inet1.URL, "GET " & sFile & " " & sFile
        
        'wait until this execution is done before going to next file
        Do
            DoEvents
        Loop Until Not Inet1.StillExecuting
    Loop
    
    iLastFTP = ftpDIR
    Inet1.Execute Inet1.URL, "DIR"

    
End Function

