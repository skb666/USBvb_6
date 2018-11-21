Attribute VB_Name = "infolParam"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Const DBT_DEVICEARRIVAL As Long = &H8000&
Public Const DBT_DEVICEREMOVECOMPLETE As Long = &H8004&
'设备类型：逻辑卷标
Public Const DBT_DEVTYP_VOLUME As Long = &H2
'与WM_DEVICECHANGE消息相关联的结构体头部信息
Public Type DEV_BROADCAST_HDR
    lSize As Long
    lDevicetype As Long    '设备类型
    lReserved As Long
End Type
'设备为逻辑卷时对应的结构体信息
Public Type DEV_BROADCAST_VOLUME
    lSize As Long
    lDevicetype As Long
    lReserved As Long
    lUnitMask As Long    '和逻辑卷标对应的掩码
    iFlag As Integer
End Type
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Info As DEV_BROADCAST_HDR
Public info_volume As DEV_BROADCAST_VOLUME
Public NewDeviceName As String
Public HaveParam As Boolean
Public OldWindowProc As Long
Public NoCopys As Boolean
Public Counter As Long
Public Const GWL_WNDPROC = (-4)
Public WM_TASKBARCREATED As Long
Public copyMin As Integer
Public copyMax As Integer
Public copyMinFileLen As Currency
Public copyMaxFileLen As Currency
Public copyMinType As Integer
Public copyMaxType As Integer
Public StopCopy As Integer
Public copyDoevents As Integer
'Download by http://www.codefans.net
Public Sub LoadSetting()
'''''''''''''''''载入相关设置
If GetIni("Setting", "copyMin", App.path & "\Setting.dat") = "1" Then copyMin = 1 Else copyMin = 0
If GetIni("Setting", "copyMax", App.path & "\Setting.dat") = "1" Then copyMax = 1 Else copyMax = 0
If GetIni("Setting", "copyDoevents", App.path & "\Setting.dat") = "1" Then copyDoevents = 1 Else copyDoevents = 0
Dim strInt As String
strInt = GetIni("Setting", "copyMinType", App.path & "\Setting.dat")
Select Case strInt
Case 0
copyMinType = 0
Case 1
copyMinType = 1
Case 2
copyMinType = 2
Case Else
copyMinType = 0
End Select
strInt = GetIni("Setting", "copyMaxType", App.path & "\Setting.dat")
Select Case strInt
Case 0
copyMaxType = 0
Case 1
copyMaxType = 1
Case 2
copyMaxType = 2
Case Else
copyMaxType = 0
End Select
If IsNumeric(GetIni("Setting", "copyMinFileLen", App.path & "\Setting.dat")) = True Then
    If Len(GetIni("Setting", "copyMinFileLen", App.path & "\Setting.dat")) > 8 Then
    copyMinFileLen = 100
    Else
    copyMinFileLen = GetIni("Setting", "copyMinFileLen", App.path & "\Setting.dat")
    End If
Else
copyMinFileLen = 100
End If
If IsNumeric(GetIni("Setting", "copyMaxFileLen", App.path & "\Setting.dat")) = True Then
    If Len(GetIni("Setting", "copyMaxFileLen", App.path & "\Setting.dat")) > 8 Then
    copyMaxFileLen = 100
    Else
    copyMaxFileLen = GetIni("Setting", "copyMaxFileLen", App.path & "\Setting.dat")
    End If
Else
copyMaxFileLen = 100
End If
''''''''''''''''''载入相关设置
End Sub

