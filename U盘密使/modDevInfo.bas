Attribute VB_Name = "modDevInfo"
Option Explicit
'Download by http://www.codefans.net
' advanced information about devices

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal cb As Long _
)

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    lpString1 As Any, lpString2 As Any _
) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" ( _
    lpString As Any _
) As Long

Private Declare Function CreateFile Lib "kernel32.dll" _
Alias "CreateFileA" ( _
    ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Private Declare Function DeviceIoControl Lib "kernel32" ( _
    ByVal hDevice As Long, ByVal dwIoControlCode As Long, _
    lpInBuffer As Any, ByVal nInBufferSize As Long, _
    lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
    lpBytesReturned As Long, lpOverlapped As Any _
) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long _
) As Long

Private Const OPEN_EXISTING                 As Long = 3&
Private Const FILE_SHARE_READ               As Long = &H1&
Private Const FILE_SHARE_WRITE              As Long = &H2&
Private Const GENERIC_READ                  As Long = &H80000000
Private Const IOCTL_STORAGE_QUERY_PROPERTY  As Long = &H2D1400

Private Type STORAGE_PROPERTY_QUERY
    PropertyId                              As STORAGE_PROPERTY_ID
    QueryType                               As STORAGE_QUERY_TYPE
    AdditionalParameters                    As Byte
End Type

Public Type DEVICE_INFORMATION
    Valid                                   As Boolean
    BusType                                 As STORAGE_BUS_TYPE
    Removable                               As Boolean
    VendorID                                As String
    ProductID                               As String
    ProductRevision                         As String
End Type

Private Type STORAGE_DEVICE_DESCRIPTOR
    Version                                 As Long
    Size                                    As Long
    DeviceType                              As Byte
    DeviceTypeModifier                      As Byte
    RemovableMedia                          As Byte
    CommandQueueing                         As Byte
    VendorIdOffset                          As Long
    ProductIdOffset                         As Long
    ProductRevisionOffset                   As Long
    SerialNumberOffset                      As Long
    BusType                                 As Integer
    RawPropertiesLength                     As Long
    RawDeviceProperties                     As Byte
End Type

Public Enum STORAGE_BUS_TYPE
    BusTypeUnknown = 0
    BusTypeScsi
    BusTypeAtapi
    BusTypeAta
    BusType1394
    BusTypeSsa
    BusTypeFibre
    BusTypeUsb
    BusTypeRAID
    BusTypeMaxReserved = &H7F
End Enum

Private Enum STORAGE_PROPERTY_ID
    StorageDeviceProperty = 0
    StorageAdapterProperty
    StorageDeviceIdProperty
End Enum

Private Enum STORAGE_QUERY_TYPE
    PropertyStandardQuery = 0
    PropertyExistsQuery
    PropertyMaskQuery
    PropertyQueryMaxDefined
End Enum
Private Declare Function GetVolumeInformation Lib "kernel32" Alias _
        "GetVolumeInformationA" (ByVal lpRootPathName As String, _
        ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize _
        As Long, lpVolumeSerialNumber As Long, _
        lpMaximumComponentLength As Long, lpFileSystemFlags As _
        Long, ByVal lpFileSystemNameBuffer As String, ByVal _
        nFileSystemNameSize As Long) As Long
Public DriveCount As Long
Public Function GetDevInfo(ByVal strDrive As String) As DEVICE_INFORMATION
    Dim hDrive          As Long
    Dim udtQuery        As STORAGE_PROPERTY_QUERY
    Dim dwOutBytes      As Long
    Dim lngResult       As Long
    Dim btBuffer(9999)  As Byte
    Dim udtOut          As STORAGE_DEVICE_DESCRIPTOR
    
    hDrive = CreateFile("\\.\" & Left$(strDrive, 1) & ":", 0, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                        ByVal 0&, OPEN_EXISTING, 0, 0)

    If hDrive = -1 Then Exit Function
    
    With udtQuery
        .PropertyId = StorageDeviceProperty
        .QueryType = PropertyStandardQuery
    End With
    
    lngResult = DeviceIoControl(hDrive, IOCTL_STORAGE_QUERY_PROPERTY, _
                                udtQuery, LenB(udtQuery), _
                                btBuffer(0), UBound(btBuffer) + 1, _
                                dwOutBytes, ByVal 0&)
        
    If lngResult Then
        CpyMem udtOut, btBuffer(0), Len(udtOut)
        
        With GetDevInfo
            .Valid = True
            .BusType = udtOut.BusType
            .Removable = CBool(udtOut.RemovableMedia)
            
            If udtOut.ProductIdOffset > 0 Then _
                .ProductID = StringCopy(VarPtr(btBuffer(udtOut.ProductIdOffset)))
            If udtOut.ProductRevisionOffset > 0 Then _
                .ProductRevision = StringCopy(VarPtr(btBuffer(udtOut.ProductRevisionOffset)))
            If udtOut.VendorIdOffset > 0 Then
                .VendorID = StringCopy(VarPtr(btBuffer(udtOut.VendorIdOffset)))
            End If
        End With
    Else
        GetDevInfo.Valid = False
    End If
    
    CloseHandle hDrive
End Function

Private Function StringCopy(ByVal pBuffer As Long) As String
    Dim tmp As String
    
    tmp = Space(lstrlen(ByVal pBuffer))
    lstrcpy ByVal tmp, ByVal pBuffer
    StringCopy = Trim$(tmp)
End Function
Public Function GetDriveName(ByVal strDrive As String) As String
    Dim sVolName As String * 256
    Dim sFileSys As String * 256
    Dim lVolSerial As Long
    Dim lMC As Long
    Dim lFileFlag As Long
    GetVolumeInformation strDrive + ":\", sVolName, 256, lVolSerial, lMC, lFileFlag, sFileSys, 256
    GetDriveName = TrimNulls(sVolName)
    If GetDriveName = "" Then
    If GetDevInfo(strDrive + ":\").Removable = True Then
    GetDriveName = "可移动磁盘"
    Else
    GetDriveName = "本地磁盘"
    End If
    End If
End Function
Public Function TrimNulls(strString As String) As String
    Dim l As Long
    l = InStr(1, strString, Chr(0))
    If l = 1 Then
        TrimNulls = ""
    ElseIf l > 0 Then
        TrimNulls = Left$(strString, l - 1)
    Else
        TrimNulls = strString
    End If
End Function

