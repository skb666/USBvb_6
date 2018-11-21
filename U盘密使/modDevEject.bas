Attribute VB_Name = "modDevEject"
'Download by http://www.codefans.net
Option Explicit

' safe ejection of devices (e.g. USB sticks)
'
' written by Daniel Aue (http://www.activevb.de/)

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long _
) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal hKey As Long _
) As Long
        
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Any _
) As Long


Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const SYNCHRONIZE               As Long = &H100000
Private Const STANDARD_RIGHTS_READ      As Long = &H20000

Private Const KEY_READ                  As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS             As Long = 0&

Private Declare Function CM_Request_Device_EjectA Lib "setupapi.dll" ( _
    ByVal hDevice As Long, lVetoType As Long, lpVetoName As Any, _
    ByVal cbVetoName As Long, ByVal dwFlags As Long _
) As Long

Private Declare Function CM_Locate_DevNodeA Lib "setupapi.dll" ( _
    hDevice As Long, lpDeviceName As Any, ByVal dwFlags As Long _
) As Long
        
Private Declare Function CM_Get_Device_IDA Lib "setupapi.dll" ( _
    ByVal hDevice As Long, ByVal lpIDBuffer As Long, _
    ByVal cbIDBuffer As Long, ByVal dwFlags As Long _
) As Long

Private Declare Function CM_Get_Device_ID_Size Lib "setupapi.dll" ( _
    ByRef lSize As Long, ByVal hDevice As Long, ByVal dwFlags As Long _
) As Long

Private Declare Function CM_Get_Parent Lib "setupapi.dll" ( _
    hParentDevice As Long, ByVal hDevice As Long, ByVal dwFlags As Long _
) As Long
        
Private Declare Function CM_Get_Child Lib "setupapi.dll" ( _
    hChildDevice As Long, ByVal hDevice As Long, ByVal dwFlags As Long _
) As Long
        
Private Declare Function CM_Get_Sibling Lib "setupapi.dll" ( _
    hSiblingDevice As Long, ByVal hDevice As Long, ByVal dwFlags As Long _
) As Long

Private Declare Function CM_Get_DevNode_Status Lib "setupapi.dll" ( _
    lStatus As Long, lProblem As Long, ByVal hDevice As Long, _
    ByVal dwFlags As Long _
) As Long

Private Const DN_REMOVABLE      As Long = &H4000
Private Const CR_SUCCESS        As Long = 0

Private Const REG_PATH_MOUNT    As String = "SYSTEM\MountedDevices"
Private Const REG_VALUE_DOSDEV  As String = "\DosDevices\"
Enum RegistryKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
    HKEY_LOCAL_MACHINE = &H80000002
End Enum
Enum RegDataTypes
REG_SZ = 1                         ' Unicode nul terminated string
REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
REG_BINARY = 3                     ' Free form binary
REG_DWORD = 4                      ' 32-bit number
REG_MULTI_SZ = 7                   ' Multiple Unicode strings
End Enum

Public Function EjectDevice(ByVal DriveLetter As String) As Boolean
    Dim strDeviceInstance   As String
    Dim btRegData()         As Byte
    Dim hDevice             As Long
    Dim lngStatus           As Long
    Dim lngProblem          As Long

    DriveLetter = UCase$(Left$(DriveLetter, 1)) & ":"
    
    If Not HKLMRegBinaryRead(REG_PATH_MOUNT, REG_VALUE_DOSDEV & DriveLetter, btRegData) Then
        Exit Function
    End If
    
    strDeviceInstance = btRegData
    If Not Left$(strDeviceInstance, 4) = "\??\" Then Exit Function
    
    strDeviceInstance = Mid$(strDeviceInstance, 5, InStr(1, strDeviceInstance, "{") - 6)
    strDeviceInstance = Replace$(strDeviceInstance, "#", "\")
    
    If CR_SUCCESS <> CM_Locate_DevNodeA(hDevice, ByVal strDeviceInstance, 0) Then
        Exit Function
    End If

    If CR_SUCCESS <> CM_Get_DevNode_Status(lngStatus, lngProblem, hDevice, 0) Then
        Exit Function
    End If
    
    Do While Not (lngStatus And DN_REMOVABLE) > 0
        If CR_SUCCESS <> CM_Get_Parent(hDevice, hDevice, 0) Then Exit Do
        If CR_SUCCESS <> CM_Get_DevNode_Status(lngStatus, lngProblem, hDevice, 0) Then Exit Do
    Loop
    
    If (lngStatus And DN_REMOVABLE) > 0 Then
        EjectDevice = CR_SUCCESS = CM_Request_Device_EjectA(hDevice, ByVal VarPtr(0), vbNullString, 0, 0)
    End If
End Function

Private Function HandleToDeviceID(hDevice As Long) As String
    Dim strDeviceID As String
    Dim cDeviceID   As Long
    
    If CM_Get_Device_ID_Size(cDeviceID, hDevice, 0) = 0 Then
        strDeviceID = Space(cDeviceID)
        
        If CM_Get_Device_IDA(hDevice, StrPtr(strDeviceID), cDeviceID, 0) > 0 Then
            strDeviceID = StrConv(strDeviceID, vbUnicode)
            strDeviceID = Left(strDeviceID, cDeviceID)
        Else
            strDeviceID = ""
        End If
    End If
    
    HandleToDeviceID = strDeviceID
End Function

Private Function HKLMRegBinaryRead(ByVal strPath As String, ByVal strValueName As String, btValue() As Byte) As Boolean
    Dim hKey        As Long
    Dim lngDataLen  As Long
    Dim lngResult   As Long
    Dim regType     As Long
    Dim btDataBuf() As Byte
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, strPath, 0, KEY_READ, hKey) = ERROR_SUCCESS Then
        If RegQueryValueEx(hKey, strValueName, 0, regType, ByVal 0&, lngDataLen) = ERROR_SUCCESS Then
            ReDim btDataBuf(lngDataLen - 1) As Byte
            If RegQueryValueEx(hKey, strValueName, 0, regType, btDataBuf(0), lngDataLen) = ERROR_SUCCESS Then
                btValue = btDataBuf
                HKLMRegBinaryRead = True
            End If
        End If
        
        RegCloseKey hKey
    End If
End Function
Public Function GetStringValue(hKey As RegistryKeys, strPath As String, strValue As String) As String
    Dim keyhand As Long, lRegResult As Long, strBuf As String, lDataBufSize As Long
    Dim lValueType As Long, lResult As Long
    
    lRegResult = RegOpenKeyEx(hKey, strPath, 0, KEY_READ, keyhand)
    lRegResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Or REG_EXPAND_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            strBuf = StripNull(strBuf)
            GetStringValue = strBuf
        End If
    End If
End Function
Private Function StripNull(ByVal WhatStr As String) As String
On Error GoTo errHandle:
    Dim pos As Integer
    
    pos = InStr(WhatStr, Chr$(0))
    If pos > 0 Then
        StripNull = Left$(WhatStr, pos - 1)
    Else
        StripNull = WhatStr
    End If
errHandle:
End Function

