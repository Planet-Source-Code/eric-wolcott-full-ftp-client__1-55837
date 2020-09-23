VERSION 5.00
Begin VB.Form frmFTP 
   Caption         =   "FTP"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   5160
   End
   Begin VB.CommandButton cmdRenameDir 
      Caption         =   "Rename Dir"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdLocalSet 
      Caption         =   "Set"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtLocalFilter 
      Height          =   285
      Left            =   1320
      TabIndex        =   28
      Text            =   "*.*"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7920
      TabIndex        =   26
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtFilter 
      Height          =   285
      Left            =   7080
      TabIndex        =   24
      Text            =   "*.*"
      Top             =   1800
      Width           =   735
   End
   Begin VB.ListBox lstFTP 
      Height          =   3375
      Left            =   5400
      TabIndex        =   23
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton cmdLocalRename 
      Caption         =   "Rename File"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdFTPRename 
      Caption         =   "Rename"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7440
      TabIndex        =   21
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdFTPDelDir 
      Caption         =   "Delete Dir"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdLocalDirDel 
      Caption         =   "Delete Dir"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdRefFTP 
      Caption         =   "Refresh"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7440
      TabIndex        =   18
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdFTPdelete 
      Caption         =   "Delete File"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdLocalDir 
      Caption         =   "Make Dir"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdFTPDir 
      Caption         =   "Make Dir"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdLocalDelete 
      Caption         =   "Delete File"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload >>"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "<< Download"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Hidden          =   -1  'True
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   3960
      TabIndex        =   9
      Top             =   480
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label label 
      Caption         =   "Loading"
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Local File Filter"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "FTP File Filter"
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "FTP Server:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ideas!!
'Options form
'Reg key remember prefs
'Log file analyzer
'Make log file dat file
'Error checking
Private Sub cmdDownload_Click()
Dim ftpname As String
Dim localname As String
Dim buffer As String
Dim buffer2 As String
buffer = Space(200)
buffer2 = Space(255)
    ftpname = lstFTP.List(lstFTP.ListIndex)
    If UCase(Dir1.Path) <> UCase("C:\") Then
        localname = Dir1.Path & "\" & ftpname
    Else
        localname = Dir1.Path & ftpname
    End If
    localname = InputBox("Enter Local File Name", "Save as", localname)
    temp = FtpGetFile(hConnection, ftpname, localname, False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0)
    File1.Refresh
    'FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, buffer, 200, ByVal 0&
    'MsgBox buffer
    FtpGetCurrentDirectory hConnection, buffer2, 255
    buffer2 = Left(buffer2, InStr(1, buffer2, " ") - 2)
    buffer2 = buffer2 & "/" & ftpname
    Print #2, Date & " " & Time & " -> Remote Filename: " & buffer2 & " DOWNLOADED to " & localname
End Sub

Private Sub cmdFTPDelDir_Click()
Dim ftpname As String
Dim buffer As String
    buffer = Space(255)
    If Left(lstFTP.List(lstFTP.ListIndex), 5) = "[DIR]" Then
        ftpname = Right(lstFTP.List(lstFTP.ListIndex), Len(lstFTP.List(lstFTP.ListIndex)) - 6)
        FtpRemoveDirectory hConnection, ftpname
        Call EnumAll(txtFilter.Text)
        FtpGetCurrentDirectory hConnection, buffer, 255
        buffer = Left(buffer, InStr(1, buffer, " ") - 2)
        buffer = buffer & "/" & ftpname
        Print #2, Date & " " & Time & " -> Remote Directory: " & buffer & " DELETED"
    Else
        MsgBox "This is not a valid Directory"
    End If
End Sub

Private Sub cmdFTPdelete_Click()
Dim buffer As String
Dim ftpname As String
    ftpname = lstFTP.List(lstFTP.ListIndex)
    If Left(ftpname, 5) = "[DIR]" Then
        MsgBox "This is not a valid file"
    Else
        buffer = Space(255)
        FtpDeleteFile hConnection, ftpname
        Call EnumAll(txtFilter.Text)
        FtpGetCurrentDirectory hConnection, buffer, 255
        buffer = Left(buffer, InStr(1, buffer, " ") - 2)
        buffer = buffer & "/" & ftpname
        Print #2, Date & " " & Time & " -> Remote Filename: " & buffer & " DELETED"
    End If
End Sub

Private Sub cmdFTPDir_Click()
Dim dirname As String
Dim buffer As String
    buffer = Space(255)
    dirname = InputBox("Enter Name of Directory", "Create Directory")
    FtpCreateDirectory hConnection, dirname
    Call EnumAll(txtFilter.Text)
    FtpGetCurrentDirectory hConnection, buffer, 255
    buffer = Left(buffer, InStr(1, buffer, " ") - 2)
    buffer = buffer & "/" & dirname
    Print #2, Date & " " & Time & " -> Remote Directory: " & buffer & " CREATED"
End Sub

Private Sub cmdFTPRename_Click()
Dim name As String
Dim buffer As String
Dim prevname
    If Left(lstFTP.List(lstFTP.ListIndex), 5) = "[DIR]" Then
        MsgBox "This is not a valid file"
    Else
        buffer = Space(255)
        prevname = lstFTP.List(lstFTP.ListIndex)
        name = InputBox("Enter New Name For " & prevname, "Rename File", prevname)
        FtpRenameFile hConnection, prevname, name
        Call EnumAll(txtFilter.Text)
        FtpGetCurrentDirectory hConnection, buffer, 255
        buffer = Left(buffer, InStr(1, buffer, " ") - 2)
        prevname = buffer & "/" & prevname
        buffer = buffer & "/" & name
        Print #2, Date & " " & Time & " -> Remote File " & prevname & " RENAMED To: " & buffer
    End If
End Sub

Private Sub cmdLocalDelete_Click()
Dim localname As String
    If UCase(Dir1.Path) <> UCase("C:\") Then
        localname = Dir1.Path & "\" & File1.List(File1.ListIndex)
    Else
        localname = Dir1.Path & File1.List(File1.ListIndex)
    End If
    DeleteFile localname
    File1.Refresh
    Print #2, Date & " " & Time & " -> Local Filename: " & localname & " DELETED"
End Sub

Private Sub cmdLocalDir_Click()
Dim pathname As String
Dim security As SECURITY_ATTRIBUTES
    pathname = InputBox("Enter Directory Name", "Create Local Directory", File1.Path)
    CreateDirectory pathname, security
    Dir1.Refresh
    Dir1.Path = pathname
    File1.Refresh
    Print #2, Date & " " & Time & " -> Local Directory: " & pathname & " CREATED"
End Sub

Private Sub cmdLocalDirDel_Click()
Dim pathname As String
    pathname = InputBox("Enter Directory Name", "Remove Directory", Dir1.List(Dir1.ListIndex))
    RemoveDirectory pathname
    Dir1.Refresh
    File1.Refresh
    Print #2, Date & " " & Time & " -> Local Directory: " & pathname & " DELETED"
End Sub

Private Sub cmdLocalRename_Click()
Dim localname As String
Dim newname As String
    If UCase(Dir1.Path) <> UCase("C:\") Then
        localname = Dir1.Path & "\" & File1.List(File1.ListIndex)
    Else
        localname = Dir1.Path & File1.List(File1.ListIndex)
    End If
    newname = InputBox("Enter New Name For File", "Rename File", localname)
    Name localname As newname
    Print #2, Date & " " & Time & " -> Local Filename: " & localname & " RENAMED to " & newname
    File1.Refresh
End Sub

Private Sub cmdLocalSet_Click()
    If Len(txtLocalFilter.Text) > 5 Or Len(txtLocalFilter.Text) < 3 Then
        MsgBox "Invalid filter"
    Else
        File1.Pattern = txtLocalFilter.Text
        File1.Refresh
    End If
End Sub

Private Sub cmdLogin_Click()
Dim sorgpath As String
Dim buffer As String
    buffer = Space(200)
    hOpen = InternetOpen("FTP", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, txtServer.Text, INTERNET_DEFAULT_FTP_PORT, txtLogin.Text, txtPass.Text, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
    InternetGetLastResponseInfo test, buffer, 200
    If InStr(1, buffer, "530") > 0 Then
        MsgBox "Invalid username or password", , "Error"
        InternetCloseHandle hOpen
    ElseIf hConnection = 0 Then
        MsgBox "Unknown Error: No connection Established", , "Error"
    ElseIf InStr(1, buffer, "230") > 0 Then
        Call EnumAll(txtFilter.Text)
        Open "C:\ftpvblog.txt" For Append As #2
        logOpen = 1
        Print #2, Date & " " & Time & " -> Session Start"
        Call enableall
    End If
End Sub

Private Sub cmdLogout_Click()
    Call disableall
    InternetCloseHandle hOpen
    InternetCloseHandle hConnection
    lstFTP.Clear
End Sub

Private Sub cmdRefFTP_Click()
    Call EnumAll(txtFilter.Text)
End Sub

Private Sub cmdRefresh_Click()
    File1.Refresh
End Sub

Private Sub cmdRenameDir_Click()
Dim name As String
Dim newname As String
    newname = InputBox("Enter New name", "Rename Dir", Dir1.List(Dir1.ListIndex))
    Name Dir1.List(Dir1.ListIndex) As newname
    Dir1.Refresh
End Sub

Private Sub cmdSet_Click()
    If Len(txtFilter.Text) > 5 Or Len(txtFilter.Text) < 3 Then
        MsgBox "Invalid filter"
    Else
        Call EnumAll(txtFilter.Text)
    End If
End Sub

Private Sub cmdUpload_Click()
Dim ftpname As String
Dim localname As String
Dim buffer As String
    buffer = Space(255)
    If UCase(Dir1.Path) <> UCase("C:\") Then
        localname = Dir1.Path & "\" & File1.List(File1.ListIndex)
    Else
        localname = Dir1.Path & File1.List(File1.ListIndex)
    End If
    ftpname = InputBox("Enter Remote File Name", "Save As", File1.List(File1.ListIndex))
    temp = FtpPutFile(hConnection, localname, ftpname, FTP_TRANSFER_TYPE_UNKNOWN, 0)
    Call EnumAll(txtFilter.Text)
    FtpGetCurrentDirectory hConnection, buffer, 255
    buffer = Left(buffer, InStr(1, buffer, " ") - 2)
    buffer = buffer & "/" & ftpname
    Print #2, Date & " " & Time & " -> Local Filename: " & localname & " UPLOADED to " & Trim(buffer)
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
    Call cmdUpload_Click
End Sub

Private Sub Form_Load()
    Dir1.Path = "C:\"
    numSub = 0
    ftpfilter = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If logOpen = 1 Then
    Print #2, Date & " " & Time & " -> Session End"
    Close #2
    logOpen = 0
End If
End Sub

Private Sub lstFTP_DblClick()
Dim fname As String
    fname = lstFTP.List(lstFTP.ListIndex)
    If Left(fname, 5) = "[DIR]" Then
        If Right(fname, Len(fname) - 6) = ".." Then
            numSub = numSub - 1
        Else
            numSub = numSub + 1
        End If
    Else
        Call cmdDownload_Click
    End If
    FtpSetCurrentDirectory hConnection, Right(fname, Len(fname) - 6)
    Call EnumAll(txtFilter.Text)
End Sub

Public Sub enableall()
    cmdRefresh.Enabled = True
    cmdRefFTP.Enabled = True
    cmdLocalDelete.Enabled = True
    cmdFTPdelete.Enabled = True
    cmdLocalDir.Enabled = True
    cmdFTPDir.Enabled = True
    cmdLocalDirDel.Enabled = True
    cmdFTPDelDir.Enabled = True
    cmdFTPRename.Enabled = True
    cmdLocalRename.Enabled = True
    cmdUpload.Enabled = True
    cmdDownload.Enabled = True
    cmdLogin.Enabled = False
    cmdLogout.Enabled = True
    cmdSet.Enabled = True
    cmdLocalSet.Enabled = True
    cmdRenameDir.Enabled = True
End Sub

Public Sub disableall()
    cmdRefresh.Enabled = False
    cmdRefFTP.Enabled = False
    cmdLocalDelete.Enabled = False
    cmdFTPdelete.Enabled = False
    cmdLocalDir.Enabled = False
    cmdFTPDir.Enabled = False
    cmdLocalDirDel.Enabled = False
    cmdFTPDelDir.Enabled = False
    cmdSet.Enabled = False
    cmdLocalSet.Enabled = False
    cmdFTPRename.Enabled = False
    cmdLocalRename.Enabled = False
    cmdUpload.Enabled = False
    cmdDownload.Enabled = False
    cmdLogin.Enabled = True
    cmdLogout.Enabled = False
    cmdRenameDir.Enabled = False
    Call Form_Unload(0)
End Sub

Private Sub Timer1_Timer()
    label.Caption = label.Caption & "."
End Sub
