Attribute VB_Name = "mdlFileSys"
Option Explicit
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH = 260
'Download by http://www.codefans.net
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Public Type SHQUERYRBINFO
    cbSize As Long
    i64Size As ULARGE_INTEGER
    i64NumItems As ULARGE_INTEGER
End Type

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Type SYSTEMTIME
    wYear             As Integer
    wMonth            As Integer
    wDayOfWeek        As Integer
    wDay              As Integer
    wHour             As Integer
    wMinute           As Integer
    wSecond           As Integer
    wMilliseconds     As Integer
End Type

Public Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Public Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public IsScan As Boolean
Private Const OFS_MAXPATHNAME = 128
Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Const OF_CREATE = &H1000
Private Const OF_WRITE = &H1
Private Const OF_READ = &H0
Private Const FILE_END = 2
Private Const FILE_BEGIN = 0
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias _
"StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef _
cchBuf As Long) As String
Private bolStop     As Boolean
Private bolReady    As Boolean

Public Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If Not FolderExists(Left$(sPath, nPos - 1)) Then
            MkDir Left$(sPath, nPos - 1)
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Not FolderExists(sPath) Then MkDir sPath
    
    CreateFolderTree = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function

Public Function GetFileFromPath(sPath As String) As String
    Dim nPos As Integer
    
    nPos = InStrRev(sPath, "\")
    
    If nPos > 0 Then
        GetFileFromPath = Mid$(sPath, nPos + 1)
    Else
        GetFileFromPath = sPath
    End If
End Function
Public Function Copys(FromPath As String, ToPath As String) As Integer
If FolderExists(ToPath) = True Or FileExists(ToPath) = True Then Copys = 2: Exit Function
If FolderExists(FromPath) = True Then
Call CopyFiles(FromPath, ToPath, 1, "")
Copys = 1
End If
If FileExists(FromPath) = True Then
 FileCopy FromPath, ToPath
 Copys = 1
End If
End Function

Public Function GetFileName(sFilename As String) As String
    Dim lPos As Long
    
    lPos = InStr(sFilename, ".")
    If lPos > 0 Then
        GetFileName = Left$(sFilename, lPos - 1)
    Else
        GetFileName = sFilename
    End If
End Function



Public Function GetFolderFromPath(sPath As String) As String
    Dim nPos As Integer
    Dim sFolder As String

    nPos = InStrRev(sPath, "\")
    
    If nPos > 0 Then
        sFolder = Left$(sPath, nPos - 1)
        If Right$(sFolder, 1) = ":" Then
            sFolder = sFolder & "\"
        End If
        GetFolderFromPath = sFolder
    End If
End Function




Public Sub ShowProperties(sFilename As String, hwndOwner As Long)
    '##############################################################################################
    'Displays the Properties of the specified file
    '##############################################################################################
    
    Dim SEI As SHELLEXECUTEINFO
    
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = hwndOwner
        .lpVerb = "properties"
        .lpFile = sFilename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    
    Call ShellExecuteEx(SEI)
End Sub
        


Public Function GetSpecialfolder(CSIDL As Long) As String
    '##############################################################################################
    'Returns the Path to a "Special" Folder (i.e. Internet History)
    '##############################################################################################
    
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function

Public Function GetFolder(hwnd As Long, Optional sPrompt As String, Optional sStartFolder As String) As String
    '##############################################################################################
    'Displays a Folder Browser to select a Folder
    '##############################################################################################
    
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim sFolder As String
    Dim pos As Integer
    
    'Fill the BROWSEINFO structure with the needed data
    With bi
        'hwnd of the window that receives messages from the call. Can be your application or the handle from GetDesktopWindow().
        .hOwner = hwnd
        
        'Pointer to the item identifier list specifying the location of the "root" folder to browse from.
        'If NULL, the desktop folder is used.
        .pidlRoot = 0&
    
        'message to be displayed in the Browse dialog
        If Len(sPrompt) = 0 Then
            .lpszTitle = "Select the folder:"
        Else
            .lpszTitle = sPrompt
        End If
    
        'the type of folder to return. - the constants perform differently for non networked pc's
        .ulFlags = &H10
        
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)
        .lParam = SHSimpleIDListFromPath(StrConv(sStartFolder, vbUnicode))
    End With
        
    'show the browse for folders dialog
     pidl = SHBrowseForFolder(bi)
    
    'the dialog has closed, so parse & display the user's returned folder selection contained in pidl
    sFolder = Space$(MAX_PATH)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal sFolder) Then
        pos = InStr(sFolder, Chr$(0))
        GetFolder = Left$(sFolder, pos - 1)
    End If
    
    Call CoTaskMemFree(pidl)
End Function




Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Select Case uMsg
    Case BFFM_INITIALIZED
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, 0&, ByVal lpData)
    Case Else
    End Select
End Function


Private Function FARPROC(pfn As Long) As Long
    '############################################################################
    'Purpose: Required by GetGolder() Function
    
    'A dummy procedure that receives and returns
    'the value of the AddressOf operator.
    
    'This workaround is needed as you can't
    'assign AddressOf directly to a member of a
    'user-defined type, but you can assign it
    'to another long and use that instead!
    '############################################################################
 
    FARPROC = pfn
End Function
Public Sub QuickCopyFile(FromPath As String, ToPath As String)
    Dim lngFrom     As Long
    Dim lngTo       As Long
    Dim c_BufSize
    c_BufSize = 100 * 1024#
    Dim myResult    As OFSTRUCT
    'Dim myOverLapped    As OVERLAPPED
    Dim lngTotal    As Long
    Dim lngCurrent  As Long
    Dim lngCopy     As Long
    Dim buf() As Byte
    Dim lCount      As Long
    Dim lBlankCount As Long
    Dim strRate     As String
    Dim lStart      As Long
    bolReady = False
    On Error Resume Next
    ReDim buf(0 To c_BufSize - 1)
    On Error GoTo CopyErr
    lngTotal = FileLen(FromPath)
    lngFrom = OpenFile(FromPath, myResult, OF_READ)
    lCount = 0
    'If myResult.nErrCode > 0 Then
    '    Err.Raise 0, , "打开源文件错误，文件:" & textFrom.Text & "  错误号:" & myResult.nErrCode
    'End If
    If Dir(ToPath) = "" Then
        lngTo = OpenFile(ToPath, myResult, OF_CREATE)
        lngCurrent = 0
    Else
        lngCurrent = FileLen(ToPath)
        lngTo = OpenFile(ToPath, myResult, OF_WRITE)
            Call SetFilePointer(lngTo, 0, 0, FILE_END)
    End If
    'If myResult.nErrCode > 0 Then
    '    Err.Raise 0, , "打开目标文件错误，文件:" & textFrom.Text & "  错误号:" & myResult.nErrCode
    'End If
    If lngCurrent >= lngTotal Then
        bolStop = True
    Else
        If lngCurrent > 0 Then
            SetFilePointer lngFrom, lngCurrent, 0, FILE_BEGIN
        End If
        bolStop = False
    End If
    lBlankCount = 0
    Do
     lCount = lCount + 1
        If bolStop = True Then GoTo CopyExit
        'picStatus.BackColor = Me.BackColor
        ReadFile lngFrom, VarPtr(buf(0)), c_BufSize, lngCopy, 0
        If lngCopy <> c_BufSize And lngCurrent <> lngTotal And lngCurrent + lngCopy <> lngTotal Then
                Exit Do
        End If
        WriteFile lngTo, VarPtr(buf(0)), lngCopy, lngCopy, 0
        lngCurrent = lngCurrent + lngCopy
        If copyDoevents = 1 Then
            DoEvents
        Else
            If lCount Mod 10 = 0 Then DoEvents
        End If
    Loop Until lngCopy <> c_BufSize
CopyExit:
    CloseHandle lngFrom
    CloseHandle lngTo
    bolReady = True
    On Error GoTo 0
    Exit Sub
CopyErr:
   'Resume
    If lngFrom <> 0 Then CloseHandle lngFrom
    If lngTo <> 0 Then CloseHandle lngTo
    bolReady = True
    On Error GoTo 0
End Sub




