Attribute VB_Name = "Files"
'****************************************************************************
'本注释由vb编程助手注释。
'vb编程助手,您编程的好帮手。
'Download by http://www.codefans.net
'发布日期：2008-7-31
'描 述：文件操作模块
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
'运行扩展API
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
'本注释由vb编程助手注释。

'vb编程助手,您编程的好帮手。

'函数：GetSysDir

'功能：获得系统system32目录

'描述：

'备注：
'****************************************************************************
Public Function GetSysDir() As String
    Dim temp As String * 256
    Dim x As Integer
    x = GetSystemDirectory(temp, Len(temp))
    GetSysDir = Left$(temp, x)
End Function
'****************************************************************************
'本注释由vb编程助手注释。

'vb编程助手,您编程的好帮手。

'函数：GetWinDir

'功能：获得Win目录

'描述：

'备注：
'****************************************************************************
Public Function GetWinDir() As String
    Dim temp As String * 256
    Dim x As Integer
    x = GetWindowsDirectory(temp, Len(temp))
    GetWinDir = Left$(temp, x)
End Function
'****************************************************************************
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：IconToPicture

'功能：ICON 转 Picture

'描述：

'备注：
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
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：GetFileIcon

'功能：获得文件图标

'描述：Filename——文件路径;SmallIcon——可选，是否以小图标显示

'备注：
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
'本注释由vb编程助手注释。

'vb编程助手,您编程的好帮手。

'函数：FileExists

'功能：判断文件是否存在

'描述：sFilename——文件路径

'备注：存在返回True，否则返回False
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
'本注释由vb编程助手注释。

'vb编程助手,您编程的好帮手。

'函数：FolderExists

'功能：判断文件夹是否存在

'描述：sFolder——文件路径

'备注：存在返回True，否则返回False
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
'本注释由vb编程助手注释。

'vb编程助手,您编程的好帮手。

'函数：SysFolder

'功能：获取各种系统文件夹路径

'描述：

'备注：
'****************************************************************************
Public Function SysFolder(ByVal asPath As String) As String
Dim s As String
Dim tmpBuffer As String
asPath = LCase(asPath)
    Select Case asPath
    Case "temp" '获取临时文件夹位置
tmpBuffer = String(255, Chr(0))
GetTempPath 256, tmpBuffer
SysFolder = Trim(Left(tmpBuffer, InStr(1, tmpBuffer, Chr(0)) - 1))
    Case "commondesktop" '获取所有用户桌面文件夹位置
SysFolder = 路径("25")
    Case "desktop" '获取当前用户桌面文件夹位置
SysFolder = 路径("0")
    Case "commonstartmenu" '获取所有用户开始菜单文件夹位置
SysFolder = 路径("22")
    Case "startmenu" '获取当前用户开始菜单文件夹位置
SysFolder = 路径("11")
    Case "commonprograms" '获取所有用户开始菜单程序文件夹位置
SysFolder = 路径("23")
    Case "programs" '获取当前用户开始菜单程序文件夹位置
SysFolder = 路径("2")
    Case "commonappdata" '获取所有用户应用程序数据文件夹位置
SysFolder = 路径("&H23&")
    Case "appdata" '获取当前用户应用程序数据文件夹位置
SysFolder = 路径("26")
    Case "commonstartup" '获取所有用户启动文件夹位置
SysFolder = 路径("24")
    Case "startup" '获取当前用户启动文件夹位置
SysFolder = 路径("&H7")
    Case "userprofile" '获取当前用户个人文件夹位置
SysFolder = Environ("USERPROFILE")
    Case "sendto" '获取当前用户发送到文件夹位置
SysFolder = 路径("9")
    Case "prg" '获取用户Program Files文件夹位置
SysFolder = 路径("&H26")
    Case "commonfavorites" '获取所有用户收藏文件夹位置
SysFolder = 路径("&H1F")
    Case "favorites" '获取当前用户收藏文件夹位置
SysFolder = 路径("&H6")
    Case "commontemplates" '获取所有用户templates文件夹位置
SysFolder = 路径("45")
    Case "templates" '获取当前用户templates文件夹位置
SysFolder = 路径("&H15")
    Case "recent" '获取当前用户Recent文件夹位置
SysFolder = 路径("&H8")
    Case "cookies" '获取当前用户cookies文件夹位置
SysFolder = 路径("&H21")
    Case "history" '获取当前用户history文件夹位置
SysFolder = 路径("&H22")
    Case "commonmanagetool" '获取所有用户管理工具文件夹位置
SysFolder = 路径("47")
    Case "managetool" '获取当前用户管理工具文件夹位置
SysFolder = 路径("&H30")
    Case "temporaryif" '获取当前用户Temporary Internet Files文件夹位置
SysFolder = 路径("&H20")
    Case "sys" '获取SYSTEM32文件夹位置
SysFolder = GetSysDir
    Case "win" '获取WINDOWS文件夹位置
SysFolder = GetWinDir
    Case "sysdir" '获取系统所在盘位置
SysFolder = Left(GetWinDir, 1) & ":"
   End Select
End Function
Private Function 路径(ByVal asPath As String) As String
Dim CSILD_NUM As Long
Dim s As String
s = asPath
CSILD_NUM = CLng(Val(s))
s = String$(MAX_PATH, 0)
SHGetSpecialFolderPath 0, s, CSILD_NUM, 0 'False
路径 = Left(s, InStr(1, s, Chr(0)) - 1)
End Function
'****************************************************************************
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：OpenAsFile

'功能：打开文件或者打开网站或者打开邮件

'描述：asPath——文件位置或者网站地址或者对方邮件地址;Line——命令行

'备注：如果是邮件地址，请在地址前加mailto: 例如：mailto:lfxwd@126.com
'****************************************************************************
Public Function OpenAsFile(asPath As String, Optional Line As String = vbNullString) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    OpenAsFile = ShellExecute(Scr_hDC, "Open", asPath, Line, "C:\", SW_SHOWNORMAL)
End Function
'****************************************************************************
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：MovePath

'功能：复制文件夹或者文件。

'描述：sPath——文件路径

'备注：成功返回True，否则返回False
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
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：KillPath

'功能：删除文件夹或者文件。

'描述：sPath——文件路径

'备注：成功返回True，否则返回False
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
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：CreateDir

'功能：创建多级目录

'描述：path——目录路径

'备注：成功返回True，否则返回False
'****************************************************************************
'功能：创建多级目录
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
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：GetApp

'功能：获取指定文件所在的目录

'描述：path——文件路径

'备注：
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
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：GetAppType

'功能：获取指定文件的文件类型,即扩展名

'描述：path——文件路径

'备注：
'****************************************************************************
Public Function GetAppType(path As String)
Dim ac As Integer
    If path = "" Then Exit Function
  ac = InStr(StrReverse(path), ".")
      If ac = 0 Then
      GetAppType = "没有文件类型"
      Exit Function
      End If
GetAppType = Right$(path, ac - 1)
End Function
'****************************************************************************
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：GetAppExe

'功能：获取指定文件的文件名

'描述：path——文件路径

'备注：
'****************************************************************************
Public Function GetAppExe(path As String)
Dim ac As Integer
    If path = "" Then Exit Function
  ac = InStr(StrReverse(path), "\")
      If ac = 0 Then
      GetAppExe = "不是有效的文件路径"
      Exit Function
      End If
GetAppExe = Right$(path, ac - 1)
End Function
'****************************************************************************
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：GetShortpath

'功能：获取指定文件路径的短文件名

'描述：File_name——文件路径

'备注：
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
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：GetLongpath

'功能：获取指定文件路径的长文件名

'描述：sShortName——短文件路径

'备注：
'****************************************************************************
Public Function GetLongpath(ByVal sShortName As String) As String
Dim sLongName As String
Dim sTemp As String
Dim iSlashPos As Integer
'在短文件名之后加上倒斜线 "\"，避免 Instr 造成错误
sShortName = sShortName & "\"
'略过磁盘代号，从第四码开始
iSlashPos = InStr(4, sShortName, "\")
'从文件名之第四码之后，一段一段处理在二个倒斜线 "\"之间的字串转换
While iSlashPos
sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
If sTemp = "" Then 'Error 52 - Bad File Name or Number
GetLongpath = ""
Exit Function
End If
sLongName = sLongName & "\" & sTemp
iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
Wend
'将转换后的文件名加上原先略过的磁盘代号，变成完整的全路径文件名
GetLongpath = Left$(sShortName, 2) & sLongName
End Function


'****************************************************************************
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：GetFileTimes

'功能：获取指定文件的时间信息

'描述：FileName——文件路径;TimesType——选择获取文件的创建时间还是修改时间

'备注：
'****************************************************************************
Public Function GetFileTimes(FileName As String, Optional TimesType As TimeType = FoundTime) As String
On Error GoTo NoFile
  Dim hwnd As Long   '文件句柄
Dim CT As FILETIME  '文件建立时间
Dim AT As FILETIME  '文件访问时间
Dim WT As FILETIME  '最后修改时间
Dim ST As SYSTEMTIME
Dim RetVal As Long  '接收返回值
    hwnd = lopen(FileName, READ_CONTROL)
    If hwnd = -1 Then
       MsgBox FileName + " 不能打开，请检查！", vbOKOnly + vbExclamation, "不能打开文件"
       Exit Function
    End If
    RetVal = GetFileTime(hwnd, CT, AT, WT)
    RetVal = FileTimeToSystemTime(CT, ST)
        Select Case TimesType
        Case FoundTime
         If ST.wHour < 16 Then
       GetFileTimes = Trim(Str(ST.wYear)) + "年" + Trim(Str(ST.wMonth)) + "月" + Trim(Str(ST.wDay)) + "日 " + Trim(Str(ST.wHour + 8)) + ":" + Trim(Str(ST.wMinute)) + ":" + Trim(Str(ST.wSecond))
       Else
       GetFileTimes = Trim(Str(ST.wYear)) + "年" + Trim(Str(ST.wMonth)) + "月" + Trim(Str(ST.wDay + 1)) + "日 " + Trim(Str(24 - ST.wHour)) + ":" + Trim(Str(ST.wMinute)) + ":" + Trim(Str(ST.wSecond))
    End If
        Case ReviseTime
    RetVal = FileTimeToSystemTime(WT, ST)
     If ST.wHour < 16 Then
       GetFileTimes = Trim(Str(ST.wYear)) + "年" + Trim(Str(ST.wMonth)) + "月" + Trim(Str(ST.wDay)) + "日 " + Trim(Str(ST.wHour + 8)) + ":" + Trim(Str(ST.wMinute)) + ":" + Trim(Str(ST.wSecond))
       Else
       GetFileTimes = Trim(Str(ST.wYear)) + "年" + Trim(Str(ST.wMonth)) + "月" + Trim(Str(ST.wDay + 1)) + "日 " + Trim(Str(24 - ST.wHour)) + ":" + Trim(Str(ST.wMinute)) + ":" + Trim(Str(ST.wSecond))
    End If
           End Select
    RetVal = CloseHandle(hwnd)  '关闭文件句柄
Exit Function
NoFile:
  Exit Function
End Function
'****************************************************************************
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：GetIni

'功能：读取ini文件信息

'描述：AppName——欲在其中查找条目的小节名称;KeyName——欲获取的项名或条目名;FileName——文件路径

'备注：
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
'本注释由VB编程助手注释。

'VB编程助手,您编程的好帮手。

'函数：CopyFiles

'功能：复制文件夹

'描述：srcPath——要复制的文件路径;dstPath——要复制到的文件路径;IncludeSubDirs——是否复制子目录;FilePat——文件过滤

'备注：
'****************************************************************************
Function CopyFiles(srcPath As String, dstPath As String, IncludeSubDirs As Integer, FilePat As String) As Integer
'这个例程演示将与文件名（包含通配符）向匹配的文件从源目录复制到目标目录
'它也可以拷贝子目录，同DOS下的XCopy/S一样
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
   '搜索所有子目录
   Do While DirReturn <> ""
      If DirReturn <> "." And DirReturn <> ".." Then
         If (GetAttr(srcPath & DirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY Then
            '如果是目录则将其添加到目录中
            dCount = dCount + 1
            d(dCount) = srcPath & DirReturn
         End If
      End If
      DirReturn = Dir
   Loop
   DirReturn = Dir(srcPath & FilePat, 0)
   '搜索所有文件
   Do While DirReturn <> ""
      If Not ((GetAttr(srcPath & DirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY) Then
         '确定在目标目录中没有同名文件，如果有，给出提示是否覆盖原文件
         On Error Resume Next
         F% = FreeFile
         Open dstPath & DirReturn For Input As #F%
         Close #F%
         If Err = 0 Then
            F% = MsgBox("文件 " & dstPath & DirReturn & " 已经存在，是否覆盖它?", 4 + 32 + 256)
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



