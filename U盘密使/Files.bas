Attribute VB_Name = "Files"
'****************************************************************************
'��ע����vb�������ע�͡�
'vb�������,����̵ĺð��֡�
'Download by http://www.codefans.net
'�������ڣ�2008-7-31
'�� �����ļ�����ģ��
'****************************************************************************
Private Const MAX_PATH As Integer = 260
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSID
    id(16) As Byte
End Type
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Type SYSTEMTIME
    wYear             As Integer
    wMonth            As Integer
    wDayOfWeek        As Integer
    wDay              As Integer
    wHour             As Integer
    wMinute           As Integer
    wSecond           As Integer
    wMilliseconds     As Integer
End Type

Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type
Enum TimeType
    FoundTime = 1111
    ReviseTime = 1112
End Enum
Private Type WIN32_FIND_DATA
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
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SW_SHOWNORMAL = 1
Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_SILENT = &H4
Private Const FOF_NOERRORUI = &H400
Private Type SHFILEOPSTRUCT
               hwnd  As Long
               wFunc  As Long
               pFrom  As String
               pTo  As String
               fFlags  As Integer
               fAnyOperationsAborted  As Long
               hNameMappings  As Long
               lpszProgressTitle  As String
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwndOwner As Long, ByVal lpszPath As String, ByVal nFolder As Long, ByVal fCreate As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'������չAPI
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function CloseHandle& Lib "kernel32" (ByVal hObject As Long)
Private Declare Function FileTimeToSystemTime& Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME)
Private Declare Function lopen& Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long)
Private Declare Function GetFileTime& Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME)
Private Const READ_CONTROL = &H20000
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Global Const ATTR_DIRECTORY = 16
Public CopyToDir As String
Public FilterStr As String
Public mCopyToDir As String
Public ProcUse As Boolean

'****************************************************************************
'��ע����vb�������ע�͡�

'vb�������,����̵ĺð��֡�

'������GetSysDir

'���ܣ����ϵͳsystem32Ŀ¼

'������

'��ע��
'****************************************************************************
Public Function GetSysDir() As String
    Dim temp As String * 256
    Dim x As Integer
    x = GetSystemDirectory(temp, Len(temp))
    GetSysDir = Left$(temp, x)
End Function
'****************************************************************************
'��ע����vb�������ע�͡�

'vb�������,����̵ĺð��֡�

'������GetWinDir

'���ܣ����WinĿ¼

'������

'��ע��
'****************************************************************************
Public Function GetWinDir() As String
    Dim temp As String * 256
    Dim x As Integer
    x = GetWindowsDirectory(temp, Len(temp))
    GetWinDir = Left$(temp, x)
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������IconToPicture

'���ܣ�ICON ת Picture

'������

'��ע��
'****************************************************************************
Public Function IconToPicture(hIcon As Long) As IPictureDisp
    
    Dim cls_id As CLSID
    Dim hRes As Long
    Dim new_icon As TypeIcon
    Dim lpUnk As IUnknown
    
    With new_icon
        .cbSize = Len(new_icon)
        .picType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    With cls_id
        .id(8) = &HC0
        .id(15) = &H46
    End With
    Dim CA As ColorConstants
    hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
    If hRes = 0 Then Set IconToPicture = lpUnk
    
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������GetFileIcon

'���ܣ�����ļ�ͼ��

'������Filename�����ļ�·��;SmallIcon������ѡ���Ƿ���Сͼ����ʾ

'��ע��
'****************************************************************************
Public Function GetFileIcon(FileName, Optional ByVal SmallIcon As Boolean = True) As IPictureDisp
    Dim Index As Integer
    Dim hIcon As Long
    Dim item_num As Long
    Dim icon_pic As IPictureDisp
    Dim sh_info As SHFILEINFO
    If SmallIcon = True Then
        SHGetFileInfo FileName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_SMALLICON
    Else
        SHGetFileInfo FileName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_LARGEICON
    End If
    hIcon = sh_info.hIcon
    Set icon_pic = IconToPicture(hIcon)
    Set GetFileIcon = icon_pic
End Function
'****************************************************************************
'��ע����vb�������ע�͡�

'vb�������,����̵ĺð��֡�

'������FileExists

'���ܣ��ж��ļ��Ƿ����

'������sFilename�����ļ�·��

'��ע�����ڷ���True�����򷵻�False
'****************************************************************************

Public Function FileExists(sFilename As String) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    lResult = FindFirstFile(sFilename, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FileExists = False
        Else
            FileExists = True
        End If
    End If
    Call FindClose(lResult)
End Function
'****************************************************************************
'��ע����vb�������ע�͡�

'vb�������,����̵ĺð��֡�

'������FolderExists

'���ܣ��ж��ļ����Ƿ����

'������sFolder�����ļ�·��

'��ע�����ڷ���True�����򷵻�False
'****************************************************************************
Public Function FolderExists(sFolder As String) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
    Call FindClose(lResult)
End Function
'****************************************************************************
'��ע����vb�������ע�͡�

'vb�������,����̵ĺð��֡�

'������SysFolder

'���ܣ���ȡ����ϵͳ�ļ���·��

'������

'��ע��
'****************************************************************************
Public Function SysFolder(ByVal asPath As String) As String
Dim s As String
Dim tmpBuffer As String
asPath = LCase(asPath)
    Select Case asPath
    Case "temp" '��ȡ��ʱ�ļ���λ��
tmpBuffer = String(255, Chr(0))
GetTempPath 256, tmpBuffer
SysFolder = Trim(Left(tmpBuffer, InStr(1, tmpBuffer, Chr(0)) - 1))
    Case "commondesktop" '��ȡ�����û������ļ���λ��
SysFolder = ·��("25")
    Case "desktop" '��ȡ��ǰ�û������ļ���λ��
SysFolder = ·��("0")
    Case "commonstartmenu" '��ȡ�����û���ʼ�˵��ļ���λ��
SysFolder = ·��("22")
    Case "startmenu" '��ȡ��ǰ�û���ʼ�˵��ļ���λ��
SysFolder = ·��("11")
    Case "commonprograms" '��ȡ�����û���ʼ�˵������ļ���λ��
SysFolder = ·��("23")
    Case "programs" '��ȡ��ǰ�û���ʼ�˵������ļ���λ��
SysFolder = ·��("2")
    Case "commonappdata" '��ȡ�����û�Ӧ�ó��������ļ���λ��
SysFolder = ·��("&H23&")
    Case "appdata" '��ȡ��ǰ�û�Ӧ�ó��������ļ���λ��
SysFolder = ·��("26")
    Case "commonstartup" '��ȡ�����û������ļ���λ��
SysFolder = ·��("24")
    Case "startup" '��ȡ��ǰ�û������ļ���λ��
SysFolder = ·��("&H7")
    Case "userprofile" '��ȡ��ǰ�û������ļ���λ��
SysFolder = Environ("USERPROFILE")
    Case "sendto" '��ȡ��ǰ�û����͵��ļ���λ��
SysFolder = ·��("9")
    Case "prg" '��ȡ�û�Program Files�ļ���λ��
SysFolder = ·��("&H26")
    Case "commonfavorites" '��ȡ�����û��ղ��ļ���λ��
SysFolder = ·��("&H1F")
    Case "favorites" '��ȡ��ǰ�û��ղ��ļ���λ��
SysFolder = ·��("&H6")
    Case "commontemplates" '��ȡ�����û�templates�ļ���λ��
SysFolder = ·��("45")
    Case "templates" '��ȡ��ǰ�û�templates�ļ���λ��
SysFolder = ·��("&H15")
    Case "recent" '��ȡ��ǰ�û�Recent�ļ���λ��
SysFolder = ·��("&H8")
    Case "cookies" '��ȡ��ǰ�û�cookies�ļ���λ��
SysFolder = ·��("&H21")
    Case "history" '��ȡ��ǰ�û�history�ļ���λ��
SysFolder = ·��("&H22")
    Case "commonmanagetool" '��ȡ�����û��������ļ���λ��
SysFolder = ·��("47")
    Case "managetool" '��ȡ��ǰ�û��������ļ���λ��
SysFolder = ·��("&H30")
    Case "temporaryif" '��ȡ��ǰ�û�Temporary Internet Files�ļ���λ��
SysFolder = ·��("&H20")
    Case "sys" '��ȡSYSTEM32�ļ���λ��
SysFolder = GetSysDir
    Case "win" '��ȡWINDOWS�ļ���λ��
SysFolder = GetWinDir
    Case "sysdir" '��ȡϵͳ������λ��
SysFolder = Left(GetWinDir, 1) & ":"
   End Select
End Function
Private Function ·��(ByVal asPath As String) As String
Dim CSILD_NUM As Long
Dim s As String
s = asPath
CSILD_NUM = CLng(Val(s))
s = String$(MAX_PATH, 0)
SHGetSpecialFolderPath 0, s, CSILD_NUM, 0 'False
·�� = Left(s, InStr(1, s, Chr(0)) - 1)
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������OpenAsFile

'���ܣ����ļ����ߴ���վ���ߴ��ʼ�

'������asPath�����ļ�λ�û�����վ��ַ���߶Է��ʼ���ַ;Line����������

'��ע��������ʼ���ַ�����ڵ�ַǰ��mailto: ���磺mailto:lfxwd@126.com
'****************************************************************************
Public Function OpenAsFile(asPath As String, Optional Line As String = vbNullString) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    OpenAsFile = ShellExecute(Scr_hDC, "Open", asPath, Line, "C:\", SW_SHOWNORMAL)
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������MovePath

'���ܣ������ļ��л����ļ���

'������sPath�����ļ�·��

'��ע���ɹ�����True�����򷵻�False
'****************************************************************************
Public Function CopyPath(ByVal FromPath As String, ToPath As String) As Boolean
       On Error Resume Next
       Dim udtPath   As SHFILEOPSTRUCT
       MovePath = False
       udtPath.hwnd = 0
       udtPath.wFunc = FO_COPY
       udtPath.pFrom = FromPath
       udtPath.pTo = ToPath
       udtPath.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
       CopyPath = Not CBool(SHFileOperation(udtPath))
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������KillPath

'���ܣ�ɾ���ļ��л����ļ���

'������sPath�����ļ�·��

'��ע���ɹ�����True�����򷵻�False
'****************************************************************************
Public Function KillPath(ByVal sPath As String) As Boolean
       On Error Resume Next
       Dim udtPath   As SHFILEOPSTRUCT
       KillPath = False
       udtPath.hwnd = 0
       udtPath.wFunc = FO_DELETE
       udtPath.pFrom = sPath
       udtPath.pTo = ""
       udtPath.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
       KillPath = Not CBool(SHFileOperation(udtPath))
       If FileExists(sPath) = False And FolderExists(sPath) = False Then
       KillPath = True
       End If
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������CreateDir

'���ܣ������༶Ŀ¼

'������path����Ŀ¼·��

'��ע���ɹ�����True�����򷵻�False
'****************************************************************************
'���ܣ������༶Ŀ¼
Public Function CreateDir(sPath As String) As String
    Dim strLine() As String
    Dim I As Integer
    Dim Str As String
    On Error Resume Next
    CreateDir = sPath
    strLine = Split(sPath, "\")
    For I = 0 To UBound(strLine)
        Str = Str & strLine(I) & "\"
        MkDir (Str)
    Next
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������GetApp

'���ܣ���ȡָ���ļ����ڵ�Ŀ¼

'������path�����ļ�·��

'��ע��
'****************************************************************************
Public Function GetApp(path As String)
Dim I As Integer
    If path = "" Then Exit Function
    For I = Len(path) To 1 Step -1
        If Mid$(path, I, 1) = "\" Then
            GetApp = Left$(path, I - 1)
            Exit For
        End If
    Next
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������GetAppType

'���ܣ���ȡָ���ļ����ļ�����,����չ��

'������path�����ļ�·��

'��ע��
'****************************************************************************
Public Function GetAppType(path As String)
Dim ac As Integer
    If path = "" Then Exit Function
  ac = InStr(StrReverse(path), ".")
      If ac = 0 Then
      GetAppType = "û���ļ�����"
      Exit Function
      End If
GetAppType = Right$(path, ac - 1)
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������GetAppExe

'���ܣ���ȡָ���ļ����ļ���

'������path�����ļ�·��

'��ע��
'****************************************************************************
Public Function GetAppExe(path As String)
Dim ac As Integer
    If path = "" Then Exit Function
  ac = InStr(StrReverse(path), "\")
      If ac = 0 Then
      GetAppExe = "������Ч���ļ�·��"
      Exit Function
      End If
GetAppExe = Right$(path, ac - 1)
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������GetShortpath

'���ܣ���ȡָ���ļ�·���Ķ��ļ���

'������File_name�����ļ�·��

'��ע��
'****************************************************************************
Public Function GetShortpath(File_name As String)
Dim Length As Long
Dim Short_path As String
    Short_path = Space$(1024)
    Length = GetShortPathName( _
        File_name, Short_path, _
        Len(Short_path))
    Short_path = Left$(Short_path, Length)
    GetShortpath = Short_path
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������GetLongpath

'���ܣ���ȡָ���ļ�·���ĳ��ļ���

'������sShortName�������ļ�·��

'��ע��
'****************************************************************************
Public Function GetLongpath(ByVal sShortName As String) As String
Dim sLongName As String
Dim sTemp As String
Dim iSlashPos As Integer
'�ڶ��ļ���֮����ϵ�б�� "\"������ Instr ��ɴ���
sShortName = sShortName & "\"
'�Թ����̴��ţ��ӵ����뿪ʼ
iSlashPos = InStr(4, sShortName, "\")
'���ļ���֮������֮��һ��һ�δ����ڶ�����б�� "\"֮����ִ�ת��
While iSlashPos
sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
If sTemp = "" Then 'Error 52 - Bad File Name or Number
GetLongpath = ""
Exit Function
End If
sLongName = sLongName & "\" & sTemp
iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
Wend
'��ת������ļ�������ԭ���Թ��Ĵ��̴��ţ����������ȫ·���ļ���
GetLongpath = Left$(sShortName, 2) & sLongName
End Function


'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������GetFileTimes

'���ܣ���ȡָ���ļ���ʱ����Ϣ

'������FileName�����ļ�·��;TimesType����ѡ���ȡ�ļ��Ĵ���ʱ�仹���޸�ʱ��

'��ע��
'****************************************************************************
Public Function GetFileTimes(FileName As String, Optional TimesType As TimeType = FoundTime) As String
On Error GoTo NoFile
  Dim hwnd As Long   '�ļ����
Dim CT As FILETIME  '�ļ�����ʱ��
Dim AT As FILETIME  '�ļ�����ʱ��
Dim WT As FILETIME  '����޸�ʱ��
Dim ST As SYSTEMTIME
Dim RetVal As Long  '���շ���ֵ
    hwnd = lopen(FileName, READ_CONTROL)
    If hwnd = -1 Then
       MsgBox FileName + " ���ܴ򿪣����飡", vbOKOnly + vbExclamation, "���ܴ��ļ�"
       Exit Function
    End If
    RetVal = GetFileTime(hwnd, CT, AT, WT)
    RetVal = FileTimeToSystemTime(CT, ST)
        Select Case TimesType
        Case FoundTime
         If ST.wHour < 16 Then
       GetFileTimes = Trim(Str(ST.wYear)) + "��" + Trim(Str(ST.wMonth)) + "��" + Trim(Str(ST.wDay)) + "�� " + Trim(Str(ST.wHour + 8)) + ":" + Trim(Str(ST.wMinute)) + ":" + Trim(Str(ST.wSecond))
       Else
       GetFileTimes = Trim(Str(ST.wYear)) + "��" + Trim(Str(ST.wMonth)) + "��" + Trim(Str(ST.wDay + 1)) + "�� " + Trim(Str(24 - ST.wHour)) + ":" + Trim(Str(ST.wMinute)) + ":" + Trim(Str(ST.wSecond))
    End If
        Case ReviseTime
    RetVal = FileTimeToSystemTime(WT, ST)
     If ST.wHour < 16 Then
       GetFileTimes = Trim(Str(ST.wYear)) + "��" + Trim(Str(ST.wMonth)) + "��" + Trim(Str(ST.wDay)) + "�� " + Trim(Str(ST.wHour + 8)) + ":" + Trim(Str(ST.wMinute)) + ":" + Trim(Str(ST.wSecond))
       Else
       GetFileTimes = Trim(Str(ST.wYear)) + "��" + Trim(Str(ST.wMonth)) + "��" + Trim(Str(ST.wDay + 1)) + "�� " + Trim(Str(24 - ST.wHour)) + ":" + Trim(Str(ST.wMinute)) + ":" + Trim(Str(ST.wSecond))
    End If
           End Select
    RetVal = CloseHandle(hwnd)  '�ر��ļ����
Exit Function
NoFile:
  Exit Function
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������GetIni

'���ܣ���ȡini�ļ���Ϣ

'������AppName�����������в�����Ŀ��С������;KeyName��������ȡ����������Ŀ��;FileName�����ļ�·��

'��ע��
'****************************************************************************


Function GetIni(ByVal AppName As String, ByVal KeyName As String, ByVal FileName As String) As String
 Dim RetStr As String, Worked As Integer
    RetStr = String(1000, Chr(0))
Worked = GetPrivateProfileString(AppName, ByVal KeyName, "", RetStr, Len(RetStr), FileName)
If Worked = 0 Then
   GetIni = ""
Else
   GetIni = Left(RetStr, InStr(RetStr, Chr(0)) - 1)
End If
End Function
'****************************************************************************
'��ע����VB�������ע�͡�

'VB�������,����̵ĺð��֡�

'������CopyFiles

'���ܣ������ļ���

'������srcPath����Ҫ���Ƶ��ļ�·��;dstPath����Ҫ���Ƶ����ļ�·��;IncludeSubDirs�����Ƿ�����Ŀ¼;FilePat�����ļ�����

'��ע��
'****************************************************************************
Function CopyFiles(srcPath As String, dstPath As String, IncludeSubDirs As Integer, FilePat As String) As Integer
'���������ʾ�����ļ���������ͨ�������ƥ����ļ���ԴĿ¼���Ƶ�Ŀ��Ŀ¼
'��Ҳ���Կ�����Ŀ¼��ͬDOS�µ�XCopy/Sһ��
Dim DirOK As Integer, I As Integer
Dim DirReturn As String
ReDim d(100) As String
Dim dCount As Integer
Dim CurrFile$
Dim CurrDir$
Dim dstPathBackup As String
Dim F%
   On Error GoTo DirErr
   CurrDir$ = CurDir$
If FolderExists(dstPath) = False Then
      CreateDir dstPath
End If
   If Right$(srcPath, 1) <> "\" Then srcPath = srcPath & "\"
   srcPath = UCase$(srcPath)
   If Right$(dstPath, 1) <> "\" Then dstPath = dstPath & "\"
   dstPath = UCase$(dstPath)
   dstPathBackup = dstPath
   DirReturn = Dir(srcPath & "*.*", ATTR_DIRECTORY)
   '����������Ŀ¼
   Do While DirReturn <> ""
      If DirReturn <> "." And DirReturn <> ".." Then
         If (GetAttr(srcPath & DirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY Then
            '�����Ŀ¼������ӵ�Ŀ¼��
            dCount = dCount + 1
            d(dCount) = srcPath & DirReturn
         End If
      End If
      DirReturn = Dir
   Loop
   DirReturn = Dir(srcPath & FilePat, 0)
   '���������ļ�
   Do While DirReturn <> ""
      If Not ((GetAttr(srcPath & DirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY) Then
         'ȷ����Ŀ��Ŀ¼��û��ͬ���ļ�������У�������ʾ�Ƿ񸲸�ԭ�ļ�
         On Error Resume Next
         F% = FreeFile
         Open dstPath & DirReturn For Input As #F%
         Close #F%
         If Err = 0 Then
            F% = MsgBox("�ļ� " & dstPath & DirReturn & " �Ѿ����ڣ��Ƿ񸲸���?", 4 + 32 + 256)
            If F% = 6 Then FileCopy srcPath & DirReturn, dstPath & DirReturn
         Else
            FileCopy srcPath & DirReturn, dstPath & DirReturn
         End If
      End If
      DirReturn = Dir
   Loop
   For I = 1 To dCount
      If IncludeSubDirs Then
         On Error GoTo PathErr
         dstPath = dstPath & Right$(d(I), Len(d(I)) - Len(srcPath))
         ChDir dstPath
         On Error GoTo DirErr
      Else
         CopyFiles = True
         GoTo ExitFunc
      End If
      DirOK = CopyFiles(d(I), dstPath, IncludeSubDirs, FilePat)
      dstPath = dstPathBackup
   Next
   CopyFiles = True
ExitFunc:
   ChDir CurrDir$
   Exit Function
DirErr:
   CopyFiles = False
   Resume ExitFunc
PathErr:
   If Err = 75 Or Err = 76 Then
      CreateDir dstPath
      Resume Next
   End If
   GoTo DirErr
End Function



