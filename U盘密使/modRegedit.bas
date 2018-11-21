Attribute VB_Name = "modRegedit"
Option Explicit

'''''''''''''''' ��ע�����в�����API
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

'����API
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'��ϵͳ��������ʱ���ؼ��ֱ�����
Private Const REG_OPTION_NON_VOLATILE = 0

'ע���ؼ��ְ�ȫѡ��...
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Private Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
'����ֵ
Private Const ERROR_NONE = 0
Private Const ERROR_BADKEY = 2
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_SUCCESS = 0

'�йص���/�����ĳ���
Private Const REG_FORCE_RESTORE As Long = 8&
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type LUID
    lowpart As Long
    highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges As LUID_AND_ATTRIBUTES
End Type

Public Enum ValueType
    REG_SZ = 1                         ' �ַ���ֵ
    REG_EXPAND_SZ = 2                  ' �������ַ���ֵ
    REG_BINARY = 3                     ' ������ֵ
    REG_DWORD = 4                      ' DWORDֵ
    REG_MULTI_SZ = 7                   ' ���ַ���ֵ
End Enum

Public Enum KeyRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private hKey As Long            ' ע������ľ��
Private i As Long               ' ѭ������
Private j As Long               ' ѭ������
Private Success As Long         ' API�����ķ���ֵ, �жϺ��������Ƿ�ɹ�

'-------------------------------------------------------------------------------------------------------------
'- �½�ע���ؼ��ֲ�����ע���ؼ��ֵ�ֵ...
'- ��� ValueName �� Value ��ȱʡ, ��ֻ�½� KeyName ����, ���Ӽ�...
'- ���ֻȱʡ ValueName ������ָ�� KeyName ��Ĭ��ֵ
'- ����˵��: KeyRoot--������, KeyName--��������, ValueName--ֵ������, Value--ֵ������, ValueType--ֵ������
'-------------------------------------------------------------------------------------------------------------
Public Function SetKeyValue(KeyRoot As KeyRoot, KeyName As String, Optional ValueName As String, Optional Value As Variant = "", Optional ValueType As ValueType = REG_SZ) As Boolean
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' ע���ȫ����
    Dim dValue(0 To 3) As Byte
    Dim sValue As String
    Dim tmpValue() As Byte
    
    lpAttr.nLength = 50                                 ' ���ð�ȫ����Ϊȱʡֵ
    lpAttr.lpSecurityDescriptor = 0                     ' ���ð�ȫ����Ϊȱʡֵ
    lpAttr.bInheritHandle = True                        ' ���ð�ȫ����Ϊȱʡֵ
    
    ' �½�ע���ؼ���
    Success = RegCreateKeyEx(KeyRoot, KeyName, 0, ValueType, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, hKey, 0)
    If Success <> ERROR_SUCCESS Then
        SetKeyValue = False
        Call RegCloseKey(hKey)
        Exit Function
    End If
    
    ' ����ע���ؼ��ֵ�ֵ...
    If IsMissing(ValueName) = False Then
        Select Case ValueType
            Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                Success = RegSetValueEx(hKey, ValueName, 0, ValueType, ByVal CStr(Value), LenB(StrConv(Value, vbFromUnicode)) + 1)
            Case REG_DWORD
                If CDbl(Value) <= 4294967295# And CDbl(Value) >= 0 Then
                    sValue = DoubleToHex(Value)
                    dValue(0) = Format("&h" & Mid(sValue, 7, 2))
                    dValue(1) = Format("&h" & Mid(sValue, 5, 2))
                    dValue(2) = Format("&h" & Mid(sValue, 3, 2))
                    dValue(3) = Format("&h" & Mid(sValue, 1, 2))
                    Success = RegSetValueEx(hKey, ValueName, 0, ValueType, dValue(0), 4)
                Else
                    Success = ERROR_BADKEY
                End If
            Case REG_BINARY
                On Error Resume Next
                Success = 1         ' �������API���ɹ�(�ɹ�����0)
                ReDim tmpValue(0 To UBound(Value))
                For i = 0 To UBound(tmpValue)
                    tmpValue(i) = Value(i)
                Next
                Success = RegSetValueEx(hKey, ValueName, 0, ValueType, tmpValue(0), UBound(Value) + 1)
        End Select
    End If
    
    If Success <> ERROR_SUCCESS Then
        SetKeyValue = False
        Call RegCloseKey(hKey)
        Exit Function
    End If
    
    Call RegCloseKey(hKey)
    SetKeyValue = True
End Function

'-------------------------------------------------------------------------------------------------------------
'- ����Ѵ��ڵ�ע���ؼ��ֵ�ֵ...
'- ��� ValueName="" �򷵻� KeyName ���Ĭ��ֵ...
'- ���ָ����ע���ؼ��ֲ�����, �򷵻ؿմ�...
'- ����˵��: KeyRoot--������, KeyName--��������, ValueName--ֵ������, ValueType--ֵ������
'-------------------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As KeyRoot, KeyName As String, ValueName As String, Optional ValueType As Long) As String
    Dim TempValue As String                             ' ע���ؼ��ֵ���ʱֵ
    Dim Value As String                                 ' ע���ؼ��ֵ�ֵ
    Dim ValueSize As Long                               ' ע���ؼ��ֵ�ֵ��ʵ�ʳ���
    Dim dValue() As Byte
    Dim bValue() As Byte
    
    TempValue = Space(1024)                             ' �洢ע���ؼ��ֵ���ʱֵ�Ļ�����
    ValueSize = 1024                                    ' ����ע���ؼ��ֵ�ֵ��Ĭ�ϳ���
    
    ' ��һ���Ѵ��ڵ�ע���ؼ���
    Call RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
    ' ����Ѵ򿪵�ע���ؼ��ֵ�ֵ
    Call RegQueryValueEx(hKey, ValueName, 0, ValueType, ByVal TempValue, ValueSize)
    
    ' ����ע���ؼ��ֵĵ�ֵ
    Select Case ValueType                                       ' ͨ���жϹؼ��ֵ�����, ���д���
        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            TempValue = Left(TempValue, ValueSize - 1)          ' ȥ��TempValueβ���ո�
            Value = TempValue
        Case REG_DWORD
            ReDim dValue(0 To 3)
            Call RegQueryValueEx(hKey, ValueName, 0, REG_DWORD, dValue(0), ValueSize)
            For i = 3 To 0 Step -1
                Value = Value + String(2 - Len(Hex(dValue(i))), "0") + Hex(dValue(i))   ' ���ɳ���Ϊ8��ʮ�������ַ���
            Next i
            If CDbl("&H" & Value) < 0 Then                                              ' ��ʮ�����Ƶ� Value ת��Ϊʮ����
                Value = 2 ^ 32 + CDbl("&H" & Value)
            Else
                Value = CDbl("&H" & Value)
            End If
        Case REG_BINARY
            If ValueSize > 0 Then
                ReDim bValue(ValueSize - 1)     ' �洢 REG_BINARY ֵ����ʱ����
                RegQueryValueEx hKey, ValueName, 0, REG_BINARY, bValue(0), ValueSize
                For i = 0 To ValueSize - 1
                    Value = Value + String(2 - Len(Hex(bValue(i))), "0") + Hex(bValue(i)) + " "  ' ������ת�����ַ���
                Next
                Erase bValue()
            End If
    End Select
    
    Call RegCloseKey(hKey)
    GetKeyValue = Trim(Value)
End Function

'-------------------------------------------------------------------------------------------------------------
'- ɾ���Ѵ��ڵ�ע���ؼ��ֵ�ֵ...
'- ���ָ����ע���ؼ��ֲ�����, �����κβ���...
'- ����˵��: KeyRoot--������, KeyName--��������, ValueName--ֵ������
'-------------------------------------------------------------------------------------------------------------
Public Function DeleteKey(KeyRoot As KeyRoot, KeyName As String, Optional ValueName As String) As Boolean
    Dim tmpKeyName As String                            ' ע���ؼ��ֵ���ʱ��������
    Dim tmpValueName As String                          ' ע���ؼ��ֵ���ʱ�Ӽ�����
    
    ' ��һ���Ѵ��ڵ�ע���ؼ���
    Success = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If Success <> ERROR_SUCCESS Then
        DeleteKey = False
        RegCloseKey hKey
        Exit Function
    End If
    
    ' ɾ���Ѵ򿪵�ע���ؼ���
    tmpKeyName = ""
    tmpValueName = KeyName
    If ValueName = "" Then                      ' �ж�ValueName�Ƿ�ȱʡ, ��ȱʡ����Ӧ����
        If InStrRev(KeyName, "\") > 1 Then
            tmpValueName = Right(KeyName, InStrRev(KeyName, "\") + 1)
            tmpKeyName = Left(KeyName, InStrRev(KeyName, "\") - 1)
        End If
        Success = RegOpenKeyEx(KeyRoot, tmpKeyName, 0, KEY_ALL_ACCESS, hKey)
        Success = RegDeleteKey(hKey, tmpValueName)
    Else
        Success = RegDeleteValue(hKey, ValueName)
    End If
    If Success <> ERROR_SUCCESS Then
        DeleteKey = False
        Call RegCloseKey(hKey)
        Exit Function
    End If
    
    ' �ر�ע���ؼ���
    Call RegCloseKey(hKey)
    DeleteKey = True
End Function

'-------------------------------------------------------------------------------------------------------------
'- ���ע���ؼ��ֵ�һЩ��Ϣ...
'- SubKeyName()      ע���ؼ��ֵ��������������(ע��:��С�±�Ϊ0)
'- ValueName()       ע���ؼ��ֵ������Ӽ�������(ע��:��С�±�Ϊ0)
'- ValueType()       ע���ؼ��ֵ������Ӽ�������(ע��:��С�±�Ϊ0)
'- CountKey          ע���ؼ��ֵ���������
'- CountValue        ע���ؼ��ֵ��Ӽ�����
'- MaxLenKey         ע���ؼ��ֵ��������Ƶ���󳤶�
'- MaxLenValue       ע���ؼ��ֵ��Ӽ����Ƶ���󳤶�
'-------------------------------------------------------------------------------------------------------------
Public Function GetKeyInfo(KeyRoot As KeyRoot, KeyName As String, SubKeyName() As String, ValueName() As String, ValueType() As ValueType, Optional CountKey As Long, Optional CountValue As Long, Optional MaxLenKey As Long, Optional MaxLenValue As Long) As Boolean
    Dim f As FILETIME
    Dim l As Long
    Dim s As String
    Dim t As ValueType
    
    ' ��һ���Ѵ��ڵ�ע���ؼ���
    Success = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If Success <> ERROR_SUCCESS Then
        GetKeyInfo = False
        Call RegCloseKey(hKey)
        Exit Function
    End If
    
    ' ���һ���Ѵ򿪵�ע���ؼ��ֵ���Ϣ
    Success = RegQueryInfoKey(hKey, vbNullString, ByVal 0&, ByVal 0&, CountKey, MaxLenKey, ByVal 0&, CountValue, MaxLenValue, ByVal 0&, ByVal 0&, f)
    If Success <> ERROR_SUCCESS Then
        GetKeyInfo = False
        Call RegCloseKey(hKey)
        Exit Function
    End If
    
    If CountKey <> 0 Then
        ReDim SubKeyName(CountKey - 1) As String            ' ���¶�������, ʹ�������С��ע���ؼ��ֵ���������ƥ��
        For i = 0 To CountKey - 1
            SubKeyName(i) = Space(255)
            l = 255
            Call RegEnumKeyEx(hKey, i, ByVal SubKeyName(i), l, 0, vbNullString, ByVal 0&, f)
            SubKeyName(i) = Left(SubKeyName(i), l)
        Next
        
        ' ����Ķ���ѭ�����ַ����������ð������
        For i = 0 To UBound(SubKeyName)
            For j = i + 1 To UBound(SubKeyName)
                If SubKeyName(i) > SubKeyName(j) Then
                    s = SubKeyName(i)
                    SubKeyName(i) = SubKeyName(j)
                    SubKeyName(j) = s
                End If
            Next j
        Next i
    End If
    
    If CountValue <> 0 Then
        ReDim ValueName(CountValue - 1) As String           ' ���¶�������, ʹ�������С��ע���ؼ��ֵ��Ӽ�����ƥ��
        ReDim ValueType(CountValue - 1) As ValueType        ' ���¶�������, ʹ�������С��ע���ؼ��ֵ��Ӽ�����ƥ��
        For i = 0 To CountValue - 1
            ValueName(i) = Space(255)
            l = 255
            Call RegEnumValue(hKey, i, ByVal ValueName(i), l, 0, ValueType(i), ByVal 0&, ByVal 0&)
            ValueName(i) = Left(ValueName(i), l)
        Next
        
        ' ����Ķ���ѭ�����ַ����������ð������
        For i = 0 To UBound(ValueName)
            For j = i + 1 To UBound(ValueName)
                If ValueName(i) > ValueName(j) Then
                    s = ValueName(i)
                    ValueName(i) = ValueName(j)
                    ValueName(j) = s
                    t = ValueType(i)
                    ValueType(i) = ValueType(j)
                    ValueType(j) = t
                End If
            Next j
        Next i
    End If
    
    ' �ر�ע���ؼ���...
    Call RegCloseKey(hKey)
    GetKeyInfo = True
End Function

'-------------------------------------------------------------------------------------------------------------
'- ����ע���ؼ��ֵ�ֵ
'- ����˵��: KeyRoot--������, KeyName--��������, FileName--�������ļ�·�����ļ���(ԭʼ���ݿ��ʽ)
'-------------------------------------------------------------------------------------------------------------
Public Function SaveKey(KeyRoot As KeyRoot, KeyName As String, FileName As String) As Boolean
    On Error Resume Next
    
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' ע���ȫ����
    lpAttr.nLength = 50                                 ' ���ð�ȫ����Ϊȱʡֵ...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...
    
    If EnablePrivilege(SE_BACKUP_NAME) = False Then
        SaveKey = False
        Exit Function
    End If
    
    Success = RegOpenKeyEx(KeyRoot, KeyName, 0&, KEY_ALL_ACCESS, hKey)
    If Success <> 0 Then
        SaveKey = False
        Success = RegCloseKey(hKey)
        Exit Function
    End If
    
    Success = RegSaveKey(hKey, FileName, lpAttr)
    If Success = 0 Then
        SaveKey = True
    Else
        SaveKey = False
    End If
    
    Success = RegCloseKey(hKey)
End Function

'-------------------------------------------------------------------------------------------------------------
'- ����ע���ؼ��ֵ�ֵ
'- ����˵��: KeyRoot--������, KeyName--��������, FileName--������ļ�·�����ļ���(ԭʼ���ݿ��ʽ)
'-------------------------------------------------------------------------------------------------------------
Public Function RestoreKey(KeyRoot As KeyRoot, KeyName As String, FileName As String) As Boolean
    On Error Resume Next
    
    If EnablePrivilege(SE_RESTORE_NAME) = False Then
        RestoreKey = False
        Exit Function
    End If
    
    Success = RegOpenKeyEx(KeyRoot, KeyName, 0&, KEY_ALL_ACCESS, hKey)
    If Success <> 0 Then
        RestoreKey = False
        Success = RegCloseKey(hKey)
        Exit Function
    End If
    
    Success = RegRestoreKey(hKey, FileName, REG_FORCE_RESTORE)
    If Success = 0 Then
        RestoreKey = True
    Else
        RestoreKey = False
    End If
    
    Success = RegCloseKey(hKey)
End Function

'-------------------------------------------------------------------------------------------------------------
'- ʹע���������/����
'-------------------------------------------------------------------------------------------------------------
Private Function EnablePrivilege(seName As String) As Boolean
    On Error Resume Next
    
    Dim p_lngRtn As Long
    Dim p_lngToken As Long
    Dim p_lngBufferLen As Long
    Dim p_typLUID As LUID
    Dim p_typTokenPriv As TOKEN_PRIVILEGES
    Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES
    
    p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
    If p_lngRtn = 0 Then
        EnablePrivilege = False
        Exit Function
    End If
    If Err.LastDllError <> 0 Then
        EnablePrivilege = False
        Exit Function
    End If
    
    p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)
    If p_lngRtn = 0 Then
      EnablePrivilege = False
      Exit Function
    End If
    
    p_typTokenPriv.PrivilegeCount = 1
    p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
    p_typTokenPriv.Privileges.pLuid = p_typLUID
    
    EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function

'-------------------------------------------------------------------------------------------------------------
'- �� Double ��( ������ 0--2^32-1 )������ת��Ϊʮ�����Ʋ���ǰ�油��
'- ����˵��: Number--Ҫת���� Double ������
'-------------------------------------------------------------------------------------------------------------
Private Function DoubleToHex(ByVal Number As Double) As String
    Dim strHex As String
    strHex = Space(8)
    For i = 1 To 8
        Select Case Number - Int(Number / 16) * 16
            Case 10
                Mid(strHex, 9 - i, 1) = "A"
            Case 11
                Mid(strHex, 9 - i, 1) = "B"
            Case 12
                Mid(strHex, 9 - i, 1) = "C"
            Case 13
                Mid(strHex, 9 - i, 1) = "D"
            Case 14
                Mid(strHex, 9 - i, 1) = "E"
            Case 15
                Mid(strHex, 9 - i, 1) = "F"
            Case Else
                Mid(strHex, 9 - i, 1) = CStr(Number - Int(Number / 16) * 16)
        End Select
        Number = Int(Number / 16)
    Next i
    DoubleToHex = strHex
End Function

