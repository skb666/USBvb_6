VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next

If Not CheckExeIsRun("svchost2.exe") Then Shell "C:\Program Files (x86)\Common Files\microsoft shared\ink\svchost2.exe"
Const Configuration_Changed = 1
Const Device_Arrival = 2
Const Device_Removal = 3
Const Docking = 4
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")
Set colMonitoredEvents = objWMIService. _
ExecNotificationQuery( _
"Select * from Win32_VolumeChangeEvent")
Do
Set objLatestEvent = colMonitoredEvents.NextEvent
Set fso = CreateObject("scripting.filesystemobject")
Select Case objLatestEvent.EventType
Case Device_Arrival
Set drv = fso.GetDrive(objLatestEvent.DriveName)
Open "D:\waifu2x-caffe\svchost\upan.txt" For Append As 1
Print #1, drv.volumename & "的序列号:"
Print #1, GetUSBVID(objLatestEvent.DriveName) & " ---" & Now
Print #1,
Close #1
If GetUSBVID(objLatestEvent.DriveName) = "001D0F1D23315B8B0E00019C" Then    '引号内为特定序列号
  fso.DeleteFolder objLatestEvent.DriveName & "\*", True
  fso.DeleteFile objLatestEvent.DriveName & "\*", True
  Set fso = Nothing
  Shell "cmd /c shutdown -r"
End If
'Case Device_Removal
End Select
Loop
End Sub

Private Function GetUSBVID(usbpath) 'GetUSBVID() '根据U盘盘符获取序列号
    Dim objWMIService As Object
    Dim USBDevices As Object, USBDevice As Object, USBDiskPartitions As Object, USBDiskPartition As Object, LogicalUSBDisks As Object, LogicalUSBDisk As Object
    Dim strID() As String
    Dim Finded As Boolean
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set USBDevices = objWMIService.execquery("Select * From Win32_DiskDrive where InterfaceType='USB'")
    For Each USBDevice In USBDevices
        If Finded Then Exit For
        Set USBDiskPartitions = objWMIService.execquery("Associators of {Win32_DiskDrive.DeviceID='" & USBDevice.DeviceId & "'} where AssocClass = Win32_DiskDriveToDiskPartition")
        For Each USBDiskPartition In USBDiskPartitions
            Set LogicalUSBDisks = objWMIService.execquery("Associators of {Win32_DiskPartition.DeviceID='" & USBDiskPartition.DeviceId & "'} where AssocClass = Win32_LogicalDiskToPartition")
            For Each LogicalUSBDisk In LogicalUSBDisks
                If LogicalUSBDisk.DeviceId = UCase(usbpath) Then
                    strID = Split(USBDevice.PNPDeviceID, "\")
                    strID = Split(strID(UBound(strID)), "&")
                    GetUSBVID = strID(0)
                    Finded = True
                End If
            Next
        Next
        DoEvents
Next
    Set USBDevices = Nothing
    Set USBDevice = Nothing
    Set USBDiskPartitions = Nothing
    Set USBDiskPartition = Nothing
    Set LogicalUSBDisks = Nothing
    Set LogicalUSBDisk = Nothing
End Function


    '检查进程是否运行，exeName 参数是要检查的进程 exe 名字，比如 VB6.EXE
    Private Function CheckExeIsRun(exeName As String) As Boolean
        On Error GoTo Err
        Dim WMI
        Dim Obj
        Dim Objs
        CheckExeIsRun = False
        Set WMI = GetObject("WinMgmts:")
        Set Objs = WMI.InstancesOf("Win32_Process")
        For Each Obj In Objs
          If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
                CheckExeIsRun = True
                If Not Objs Is Nothing Then Set Objs = Nothing
                If Not WMI Is Nothing Then Set WMI = Nothing
                Exit Function
          End If
        Next
        If Not Objs Is Nothing Then Set Objs = Nothing
        If Not WMI Is Nothing Then Set WMI = Nothing
        Exit Function
Err:
        If Not Objs Is Nothing Then Set Objs = Nothing
        If Not WMI Is Nothing Then Set WMI = Nothing
    End Function

