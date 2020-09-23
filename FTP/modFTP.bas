Attribute VB_Name = "modFTP"
Public Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const INTERNET_DEFAULT_FTP_PORT = 21               ' default for FTP servers
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_FLAG_PASSIVE = &H8000000            ' used for FTP connections
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0                    ' use registry configuration
Public Const INTERNET_OPEN_TYPE_DIRECT = 1                        ' direct to net
Public Const INTERNET_OPEN_TYPE_PROXY = 3                         ' via named proxy
Public Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4   ' prevent using java/script/INS
Public Const MAX_PATH = 260
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const LANG_NEUTRAL = &H0
Public Const SUBLANG_DEFAULT = &H1
Public Const ERROR_BAD_USERNAME = 2202&
Public hConnection As Long
Public hOpen As Long
Public numSub As Integer
Public logOpen As Integer
Public ftpfilter As Integer
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Public Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Sub EnumDir()
    Dim pData As WIN32_FIND_DATA, hFind As Long, lRet As Long
    pData.cFileName = String(MAX_PATH, 0)
    hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
    If hFind = 0 Then Exit Sub
    If pData.nFileSizeLow = 0 Then
        frmFTP.lstFTP.AddItem "[DIR]" & Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    End If
    Do
        pData.cFileName = String(MAX_PATH, 0)
        lRet = InternetFindNextFile(hFind, pData)
        If lRet = 0 Then Exit Do
        If pData.nFileSizeLow = 0 Then
            frmFTP.lstFTP.AddItem "[DIR] " & Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        End If
    Loop
    InternetCloseHandle hFind
End Sub

Public Sub EnumFiles(ftpfilter As String)
    Dim pData As WIN32_FIND_DATA, hFind As Long, lRet As Long
    pData.cFileName = String(MAX_PATH, 0)
    hFind = FtpFindFirstFile(hConnection, ftpfilter, pData, 0, 0)
    If hFind = 0 Then Exit Sub
    If pData.nFileSizeLow > 0 Then
        frmFTP.lstFTP.AddItem Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    End If
    Do
        pData.cFileName = String(MAX_PATH, 0)
        lRet = InternetFindNextFile(hFind, pData)
        If lRet = 0 Then Exit Do
        If pData.nFileSizeLow > 0 Then
            frmFTP.lstFTP.AddItem Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        End If
    Loop
    InternetCloseHandle hFind
End Sub

Public Sub EnumAll(ftpfilter As String)
    frmFTP.lstFTP.Clear
    If numSub > 0 Then
        frmFTP.lstFTP.AddItem "[DIR] .."
    End If
    Call EnumDir
    Call EnumFiles(ftpfilter)
End Sub

