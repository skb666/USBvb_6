VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetUSBPath(usbid As String) As String '根据U盘序列号获取盘符
    Dim objWMIService As Object
    Dim USBDevices As Object, USBDevice As Object, USBDiskPartitions As Object, USBDiskPartition As Object, LogicalUSBDisks As Object, LogicalUSBDisk As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set USBDevices = objWMIService.execquery("Select * From Win32_DiskDrive where InterfaceType='USB'")
    For Each USBDevice In USBDevices
        If USBDevice.PNPDeviceID Like "*\" & usbid & "&*" Then
            Set USBDiskPartitions = objWMIService.execquery("Associators of {Win32_DiskDrive.DeviceID='" & USBDevice.DeviceId & "'} where AssocClass = Win32_DiskDriveToDiskPartition")
            For Each USBDiskPartition In USBDiskPartitions
                Set LogicalUSBDisks = objWMIService.execquery("Associators of {Win32_DiskPartition.DeviceID='" & USBDiskPartition.DeviceId & "'} where AssocClass = Win32_LogicalDiskToPartition")
                For Each LogicalUSBDisk In LogicalUSBDisks
                    GetUSBPath = LogicalUSBDisk.DeviceId
                Next
            Next
            Exit For
        End If
    Next
    Set USBDevices = Nothing
    Set USBDevice = Nothing
    Set USBDiskPartitions = Nothing
    Set USBDiskPartition = Nothing
    Set LogicalUSBDisks = Nothing
    Set LogicalUSBDisk = Nothing
End Function

Private Function GetUSBVID(usbPath As String) As String '根据U盘盘符获取序列号
    Dim objWMIService As Object
    Dim USBDevices As Object, USBDevice As Object, USBDiskPartitions As Object, USBDiskPartition As Object, LogicalUSBDisks As Object, LogicalUSBDisk As Object
    Dim strID() As String
    Dim Finded As Boolean
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set USBDevices = objWMIService.execquery("Select * From Win32_DiskDrive where InterfaceType='USB'")
    For Each USBDevice In USBDevices
        List1.AddItem USBDevice.PNPDeviceID
        If Finded Then Exit For
        Set USBDiskPartitions = objWMIService.execquery("Associators of {Win32_DiskDrive.DeviceID='" & USBDevice.DeviceId & "'} where AssocClass = Win32_DiskDriveToDiskPartition")
        For Each USBDiskPartition In USBDiskPartitions
            Set LogicalUSBDisks = objWMIService.execquery("Associators of {Win32_DiskPartition.DeviceID='" & USBDiskPartition.DeviceId & "'} where AssocClass = Win32_LogicalDiskToPartition")
            For Each LogicalUSBDisk In LogicalUSBDisks
                If LogicalUSBDisk.DeviceId = UCase(usbPath) Then
                    strID = Split(USBDevice.PNPDeviceID, "\")
                    strID = Split(strID(UBound(strID)), "&")
                    GetUSBVID = strID(0)
                    Finded = True
                End If
            Next
        Next
    Next
    Set USBDevices = Nothing
    Set USBDevice = Nothing
    Set USBDiskPartitions = Nothing
    Set USBDiskPartition = Nothing
    Set LogicalUSBDisks = Nothing
    Set LogicalUSBDisk = Nothing
End Function

Private Sub Command1_Click()
    MsgBox GetUSBPath("U盘序列号")
End Sub

Private Sub Command2_Click()
    MsgBox GetUSBVID("U盘盘符")
End Sub
